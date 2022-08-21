using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Collections.Immutable;
using System.Drawing;
using System.IO.Compression;
using System.Net;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class StylesXml
{
    // One built-in, and one for the default DateTime number format.
    public const int DefaultStyleCount = 2;

    private const string Header =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

    private const string XmlPart1 =
        "<borders count=\"1\">" +
        "<border>" +
        "<left/>" +
        "<right/>" +
        "<top/>" +
        "<bottom/>" +
        "<diagonal/>" +
        "</border>" +
        "</borders>" +
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

    private const int CustomNumberFormatStartId = 166;
    private static ImmutableDictionary<string, int> DefaultNumberFormatDictionary { get; } = CreateDefaultNumberFormatDictionary();
    private static ImmutableDictionary<string, int> CreateDefaultNumberFormatDictionary()
    {
        var builder = ImmutableDictionary.CreateBuilder<string, int>(StringComparer.Ordinal);
        builder.Add(NumberFormats.DateTimeUniversalSortable, 165);
        return builder.ToImmutable();
    }

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        ICollection<ImmutableStyle> styles,
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
        ICollection<ImmutableStyle> styles,
        CancellationToken token)
    {
        buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

        var customNumberFormatLookup = await WriteNumberFormatsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var fontLookup = await WriteFontsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var fillLookup = await WriteFillsAsync(stream, buffer, styles, token).ConfigureAwait(false);

        await buffer.WriteAsciiStringAsync(XmlPart1, stream, token).ConfigureAwait(false);

        var styleCount = styles.Count + DefaultStyleCount;
        if (styleCount.GetNumberOfDigits() > buffer.FreeCapacity)
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        var sb = new StringBuilder();
        sb.Append(styleCount).Append("\">");

        // The default styles must come before any other style.
        sb.AppendDefaultStyles(customNumberFormatLookup);
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        foreach (var style in styles)
        {
            sb.Clear();

            var numberFormatId = GetNumberFormatId(style.NumberFormat, customNumberFormatLookup);
            sb.Append("<xf numFmtId=\"").Append(numberFormatId).Append('"');
            if (numberFormatId > 0) sb.Append(" applyNumberFormat=\"1\"");

            var fontIndex = fontLookup[style.Font];
            sb.Append(" fontId=\"").Append(fontIndex).Append('"');
            if (fontIndex > 0) sb.Append(" applyFont=\"1\"");

            var fillIndex = fillLookup[style.Fill];
            sb.Append(" fillId=\"").Append(fillIndex).Append('"');
            if (fillIndex > 1) sb.Append(" applyFill=\"1\"");

            sb.Append(" xfId=\"0\"/>");
            await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);
        }

        await buffer.WriteAsciiStringAsync(XmlPart2, stream, token).ConfigureAwait(false);
        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
    }

    private static void AppendDefaultStyles(this StringBuilder sb, IReadOnlyDictionary<string, int> customNumberFormats)
    {
        // Index 0: The built-in default style must be the first one (meaning the first <xf> element).
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\"/>");

        // Index 1: The default style for DateTime. Expected to be index 1 by the DateTime cell value writer.
        var dateTimeNumberFormatId = GetNumberFormatId(NumberFormats.DateTimeUniversalSortable, customNumberFormats);
        sb.Append("<xf numFmtId=\"")
            .Append(dateTimeNumberFormatId)
            .Append("\" applyNumberFormat=\"1\" fontId=\"0\" fillId=\"0\" xfId=\"0\"/>");
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
        ICollection<ImmutableStyle> styles,
        CancellationToken token)
    {
        var dict = CreateCustomNumberFormatDictionary(styles);
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

    private static IReadOnlyDictionary<string, int> CreateCustomNumberFormatDictionary(ICollection<ImmutableStyle> styles)
    {
        if (styles.Count == 0)
            return DefaultNumberFormatDictionary;

        var numberFormatId = CustomNumberFormatStartId;
        var dictionary = new Dictionary<string, int>(DefaultNumberFormatDictionary, StringComparer.Ordinal);

        foreach (var style in styles)
        {
            var numberFormat = style.NumberFormat;
            if (numberFormat is null) continue;
            if (NumberFormats.GetPredefinedNumberFormatId(numberFormat) is not null) continue;

            if (dictionary.TryAdd(numberFormat, numberFormatId))
                ++numberFormatId;
        }

        return dictionary;
    }

    private static async ValueTask<Dictionary<ImmutableFont, int>> WriteFontsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        ICollection<ImmutableStyle> styles,
        CancellationToken token)
    {
        var defaultFont = new ImmutableFont();
        const int defaultCount = 1;

        var uniqueFonts = new Dictionary<ImmutableFont, int> { { defaultFont, 0 } };
        foreach (var style in styles)
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
        ICollection<ImmutableStyle> styles,
        CancellationToken token)
    {
        var defaultFill = new ImmutableFill();
        const int defaultCount = 2;

        var uniqueFills = new Dictionary<ImmutableFill, int> { { defaultFill, 0 } };
        foreach (var style in styles)
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

    private static string HexString(Color c) => $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
}
