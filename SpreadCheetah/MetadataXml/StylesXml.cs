using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO.Compression;
using System.Net;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class StylesXml
{
    private const string Header =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

    private const string XmlPart1 =
        "<cellStyleXfs count=\"1\">" +
        "<xf numFmtId=\"0\" fontId=\"0\"/>" +
        "</cellStyleXfs>" +
        "<cellXfs count=\"";

    private const string XmlPart2 =
        "</cellXfs>" +
        "<cellStyles count=\"1\">" +
        "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" +
        "</cellStyles>" +
        "<dxfs count=\"0\"/>" +
        "</styleSheet>";

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("xl/styles.xml", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            await WriteAsync(stream, buffer, styles, token).ConfigureAwait(false);
        }
    }

    private static async ValueTask WriteAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

        var customNumberFormatLookup = await WriteNumberFormatsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var fontLookup = await WriteFontsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var fillLookup = await WriteFillsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var borderLookup = await WriteBordersAsync(stream, buffer, styles, token).ConfigureAwait(false);

        await buffer.WriteAsciiStringAsync(XmlPart1, stream, token).ConfigureAwait(false);

        // Must at least have the built-in default style
        var styleCount = styles.Count;
        if (styleCount.GetNumberOfDigits() > buffer.FreeCapacity)
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        var sb = new StringBuilder();
        sb.Append(styleCount).Append("\">");

        // Assuming here that the built-in default style is part of 'styles'
        foreach (var style in styles.Keys)
        {
            sb.AppendStyle(style, customNumberFormatLookup, fontLookup, fillLookup, borderLookup);
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);
            sb.Clear();
        }

        await buffer.WriteAsciiStringAsync(XmlPart2, stream, token).ConfigureAwait(false);
        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
    }

    private static void AppendStyle(this StringBuilder sb,
        in ImmutableStyle style,
        IReadOnlyDictionary<string, int> customNumberFormats,
        IReadOnlyDictionary<ImmutableFont, int> fonts,
        IReadOnlyDictionary<ImmutableFill, int> fills,
        IReadOnlyDictionary<ImmutableBorder, int> borders)
    {
        var numberFormatId = GetNumberFormatId(style.NumberFormat, customNumberFormats);
        sb.Append("<xf numFmtId=\"").Append(numberFormatId).Append('"');
        if (numberFormatId > 0) sb.Append(" applyNumberFormat=\"1\"");

        var fontIndex = fonts[style.Font];
        sb.Append(" fontId=\"").Append(fontIndex).Append('"');
        if (fontIndex > 0) sb.Append(" applyFont=\"1\"");

        var fillIndex = fills[style.Fill];
        sb.Append(" fillId=\"").Append(fillIndex).Append('"');
        if (fillIndex > 1) sb.Append(" applyFill=\"1\"");

        if (borders.TryGetValue(style.Border, out var borderIndex) && borderIndex > 0)
            sb.Append(" borderId=\"").Append(borderIndex).Append("\" applyBorder=\"1\"");

        sb.Append(" xfId=\"0\"");

        var defaultAlignment = new ImmutableAlignment();
        if (style.Alignment == defaultAlignment)
            sb.Append("/>");
        else
            sb.Append(" applyAlignment=\"1\">").AppendAlignment(style.Alignment).Append("</xf>");
    }

    private static int GetNumberFormatId(string? numberFormat, IReadOnlyDictionary<string, int> customNumberFormats)
    {
        if (numberFormat is null) return 0;
        return NumberFormats.GetPredefinedNumberFormatId(numberFormat)
            ?? customNumberFormats.GetValueOrDefault(numberFormat);
    }

    private static async ValueTask<IReadOnlyDictionary<string, int>> WriteNumberFormatsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        if (!TryCreateCustomNumberFormatDictionary(styles, out var dict))
        {
            await buffer.WriteAsciiStringAsync("<numFmts count=\"0\"/>", stream, token).ConfigureAwait(false);
            return ImmutableDictionary<string, int>.Empty;
        }

        var sb = new StringBuilder();
        sb.Append("<numFmts count=\"").Append(dict.Count).Append("\">");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        foreach (var numberFormat in dict)
        {
            sb.Clear();
            sb.AppendNumberFormat(numberFormat.Key, numberFormat.Value);
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);
        }

        await buffer.WriteAsciiStringAsync("</numFmts>", stream, token).ConfigureAwait(false);
        return dict;
    }

    private static bool TryCreateCustomNumberFormatDictionary(
        Dictionary<ImmutableStyle, int> styles,
        [NotNullWhen(true)] out Dictionary<string, int>? dictionary)
    {
        dictionary = null;
        var numberFormatId = 165; // Custom formats start sequentially from this ID

        foreach (var style in styles.Keys)
        {
            var numberFormat = style.NumberFormat;
            if (numberFormat is null) continue;
            if (NumberFormats.GetPredefinedNumberFormatId(numberFormat) is not null) continue;

            dictionary ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (dictionary.TryAdd(numberFormat, numberFormatId))
                ++numberFormatId;
        }

        return dictionary is not null;
    }

    private static async ValueTask<Dictionary<ImmutableFont, int>> WriteFontsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var defaultFont = new ImmutableFont();
        const int defaultCount = 1;

        var uniqueFonts = new Dictionary<ImmutableFont, int> { { defaultFont, 0 } };
        foreach (var style in styles.Keys)
        {
            var font = style.Font;
            uniqueFonts[font] = 0;
        }

        var sb = new StringBuilder();
        var totalCount = uniqueFonts.Count + defaultCount - 1;
        sb.Append("<fonts count=\"").Append(totalCount).Append("\">");

        // The default font must be the first one (index 0)
        sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        var fontIndex = defaultCount;
#if NET5_0_OR_GREATER // https://github.com/dotnet/runtime/issues/34606
        foreach (var font in uniqueFonts.Keys)
#else
        foreach (var font in uniqueFonts.Keys.ToArray())
#endif
        {
            if (font.Equals(defaultFont)) continue;

            sb.Clear();
            sb.AppendFont(font);
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

            uniqueFonts[font] = fontIndex;
            ++fontIndex;
        }

        await buffer.WriteAsciiStringAsync("</fonts>", stream, token).ConfigureAwait(false);
        return uniqueFonts;
    }

    private static async ValueTask<Dictionary<ImmutableFill, int>> WriteFillsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var defaultFill = new ImmutableFill();
        const int defaultCount = 2;

        var uniqueFills = new Dictionary<ImmutableFill, int> { { defaultFill, 0 } };
        foreach (var style in styles.Keys)
        {
            var fill = style.Fill;
            uniqueFills[fill] = 0;
        }

        var sb = new StringBuilder();
        var totalCount = uniqueFills.Count + defaultCount - 1;
        sb.Append("<fills count=\"").Append(totalCount).Append("\">");

        // The 2 default fills must come first
        sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
        sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        var fillIndex = defaultCount;
#if NET5_0_OR_GREATER // https://github.com/dotnet/runtime/issues/34606
        foreach (var fill in uniqueFills.Keys)
#else
        foreach (var fill in uniqueFills.Keys.ToArray())
#endif
        {
            if (fill.Equals(defaultFill)) continue;

            sb.Clear();
            sb.AppendFill(fill);
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

            uniqueFills[fill] = fillIndex;
            ++fillIndex;
        }

        await buffer.WriteAsciiStringAsync("</fills>", stream, token).ConfigureAwait(false);
        return uniqueFills;
    }

    private static async ValueTask<Dictionary<ImmutableBorder, int>> WriteBordersAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var defaultBorder = new ImmutableBorder();
        const int defaultCount = 1;

        var uniqueBorders = new Dictionary<ImmutableBorder, int> { { defaultBorder, 0 } };
        foreach (var style in styles.Keys)
        {
            var border = style.Border;
            uniqueBorders[border] = 0;
        }

        var sb = new StringBuilder();
        var totalCount = uniqueBorders.Count + defaultCount - 1;
        sb.Append("<borders count=\"").Append(totalCount).Append("\">");

        // The default border must be the first one (index 0)
        sb.Append("<border><left/><right/><top/><bottom/><diagonal/></border>");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        var borderIndex = defaultCount;
#if NET5_0_OR_GREATER // https://github.com/dotnet/runtime/issues/34606
        foreach (var border in uniqueBorders.Keys)
#else
        foreach (var border in uniqueBorders.Keys.ToArray())
#endif
        {
            if (border.Equals(defaultBorder)) continue;

            sb.Clear();
            sb.AppendBorder(border);
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

            uniqueBorders[border] = borderIndex;
            ++borderIndex;
        }

        await buffer.WriteAsciiStringAsync("</borders>", stream, token).ConfigureAwait(false);
        return uniqueBorders;
    }

    private static void AppendNumberFormat(this StringBuilder sb, string numberFormat, int id)
    {
        sb.Append("<numFmt numFmtId=\"")
            .Append(id)
            .Append("\" formatCode=\"")
            .Append(WebUtility.HtmlEncode(numberFormat))
            .Append("\"/>");
    }

    private static void AppendFont(this StringBuilder sb, ImmutableFont font)
    {
        sb.Append("<font>");

        if (font.Bold) sb.Append("<b/>");
        if (font.Italic) sb.Append("<i/>");
        if (font.Strikethrough) sb.Append("<strike/>");

        sb.Append("<sz val=\"").AppendDouble(font.Size).Append("\"/>");

        if (font.Color is not null)
            sb.Append("<color rgb=\"").Append(HexString(font.Color.Value)).Append("\"/>");

        var fontName = font.Name ?? "Calibri";
        sb.Append("<name val=\"").Append(fontName).Append("\"/></font>");
    }

    private static void AppendFill(this StringBuilder sb, ImmutableFill fill)
    {
        if (fill.Color is null) return;
        sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"");
        sb.Append(HexString(fill.Color.Value));
        sb.Append("\"/></patternFill></fill>");
    }

    private static void AppendBorder(this StringBuilder sb, in ImmutableBorder border)
    {
        sb.Append("<border");

        if (border.Diagonal.Type.HasFlag(DiagonalBorderType.DiagonalUp))
            sb.Append(" diagonalUp=\"true\"");
        if (border.Diagonal.Type.HasFlag(DiagonalBorderType.DiagonalDown))
            sb.Append(" diagonalDown=\"true\"");

        sb.Append('>');
        sb.AppendEdgeBorder("left", border.Left);
        sb.AppendEdgeBorder("right", border.Right);
        sb.AppendEdgeBorder("top", border.Top);
        sb.AppendEdgeBorder("bottom", border.Bottom);
        sb.AppendBorderPart("diagonal", border.Diagonal.BorderStyle, border.Diagonal.Color);

        sb.Append("</border>");
    }

    private static void AppendEdgeBorder(this StringBuilder sb, string borderName, ImmutableEdgeBorder border)
        => sb.AppendBorderPart(borderName, border.BorderStyle, border.Color);

    private static void AppendBorderPart(this StringBuilder sb, string borderName, BorderStyle style, Color? color)
    {
        sb.Append('<').Append(borderName);

        if (style == BorderStyle.None)
        {
            sb.Append("/>");
            return;
        }

        sb.Append(" style=\"").AppendStyleAttributeValue(style);

        if (color is null)
        {
            sb.Append("\"/>");
            return;
        }

        sb.Append("\"><color rgb=\"")
            .Append(HexString(color.Value))
            .Append("\"/></")
            .Append(borderName)
            .Append('>');
    }

    private static void AppendStyleAttributeValue(this StringBuilder sb, BorderStyle style)
    {
        var value = style switch
        {
            BorderStyle.DashDot => "dashDot",
            BorderStyle.DashDotDot => "dashDotDot",
            BorderStyle.Dashed => "dashed",
            BorderStyle.Dotted => "dotted",
            BorderStyle.DoubleLine => "double",
            BorderStyle.Hair => "hair",
            BorderStyle.Medium => "medium",
            BorderStyle.MediumDashDot => "mediumDashDot",
            BorderStyle.MediumDashDotDot => "mediumDashDotDot",
            BorderStyle.MediumDashed => "mediumDashed",
            BorderStyle.SlantDashDot => "slantDashDot",
            BorderStyle.Thick => "thick",
            BorderStyle.Thin => "thin",
            _ => ""
        };

        sb.Append(value);
    }

    private static StringBuilder AppendAlignment(this StringBuilder sb, ImmutableAlignment alignment)
    {
        sb.Append("<alignment")
            .AppendHorizontalAlignment(alignment.Horizontal)
            .AppendVerticalAlignment(alignment.Vertical);

        if (alignment.WrapText)
            sb.Append(" wrapText=\"1\"");

        if (alignment.Indent > 0)
            sb.Append(" indent=\"").Append(alignment.Indent).Append('"');

        return sb.Append("/>");
    }

    private static StringBuilder AppendHorizontalAlignment(this StringBuilder sb, HorizontalAlignment alignment) => alignment switch
    {
        HorizontalAlignment.Left => sb.Append(" horizontal=\"left\""),
        HorizontalAlignment.Center => sb.Append(" horizontal=\"center\""),
        HorizontalAlignment.Right => sb.Append(" horizontal=\"right\""),
        _ => sb
    };

    private static StringBuilder AppendVerticalAlignment(this StringBuilder sb, VerticalAlignment alignment) => alignment switch
    {
        VerticalAlignment.Center => sb.Append(" vertical=\"center\""),
        VerticalAlignment.Top => sb.Append(" vertical=\"top\""),
        _ => sb
    };

    private static string HexString(Color c) => $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
}
