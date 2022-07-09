using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Drawing;
using System.IO.Compression;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class StylesXml
{
    private const string Header =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<numFmts count=\"0\"/>";

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

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<Style> styles,
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
        List<Style> styles,
        CancellationToken token)
    {
        buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

        var fontLookup = await WriteFontsAsync(stream, buffer, styles, token).ConfigureAwait(false);
        var fillLookup = await WriteFillsAsync(stream, buffer, styles, token).ConfigureAwait(false);

        await buffer.WriteAsciiStringAsync(XmlPart1, stream, token).ConfigureAwait(false);

        var styleCount = styles.Count + 1;
        if (styleCount.GetNumberOfDigits() > buffer.FreeCapacity)
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        var sb = new StringBuilder();
        sb.Append(styleCount);

        // The default style must be the first one (index 0)
        sb.Append("\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\"/>");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        foreach (var style in styles)
        {
            sb.Clear();

            var numberFormatId = GetNumberFormatId(style.NumberFormat);
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

    private static int GetNumberFormatId(string? numberFormat)
    {
        if (numberFormat is null) return 0;
        // TODO: Handle custom formats
        // TODO: Custom formats must be XML encoded
        // TODO: Start with numFmtId = 165 for custom formats
        return NumberFormats.GetBuiltInNumberFormatId(numberFormat) ?? 0;
    }

    private static async ValueTask<Dictionary<Font, int>> WriteFontsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        List<Style> styles,
        CancellationToken token)
    {
        var defaultFont = new Font();
        const int defaultCount = 1;

        var uniqueFonts = new Dictionary<Font, int> { { defaultFont, 0 } };
        for (var i = 0; i < styles.Count; ++i)
        {
            var font = styles[i].Font;
            uniqueFonts[font] = 0;
        }

        var sb = new StringBuilder();
        var totalCount = uniqueFonts.Count + defaultCount - 1;
        sb.Append("<fonts count=\"").Append(totalCount).Append("\">");

        // The default font must be the first one (index 0)
        sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
        await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);

        var fontIndex = defaultCount;
        foreach (var font in uniqueFonts.Keys.ToArray())
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

    private static async ValueTask<Dictionary<Fill, int>> WriteFillsAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        List<Style> styles,
        CancellationToken token)
    {
        var defaultFill = new Fill();
        const int defaultCount = 2;

        var uniqueFills = new Dictionary<Fill, int> { { defaultFill, 0 } };
        for (var i = 0; i < styles.Count; ++i)
        {
            var fill = styles[i].Fill;
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
        foreach (var fill in uniqueFills.Keys.ToArray())
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

    private static void AppendFont(this StringBuilder sb, Font font)
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

    private static void AppendFill(this StringBuilder sb, Fill fill)
    {
        if (fill.Color is null) return;
        sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"");
        sb.Append(HexString(fill.Color.Value));
        sb.Append("\"/></patternFill></fill>");
    }

    private static string HexString(Color c) => $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
}
