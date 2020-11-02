using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Collections.Generic;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class StylesXml
    {
        private const string Header =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<numFmts count=\"0\" />";

        private const string XmlPart1 =
            "</fonts>" +
            "<fills count=\"2\">" +
            "<fill><patternFill patternType=\"none\" /></fill>" +
            "<fill><patternFill patternType=\"gray125\" /></fill>" +
            "</fills>" +
            "<borders count=\"1\">" +
            "<border>" +
            "<left />" +
            "<right />" +
            "<top />" +
            "<bottom />" +
            "<diagonal />" +
            "</border>" +
            "</borders>" +
            "<cellStyleXfs count=\"1\">" +
            "<xf numFmtId=\"0\" fontId=\"0\" />" +
            "</cellStyleXfs>" +
            "<cellXfs count=\"";

        private const string XmlPart2 =
            "</cellXfs>" +
            "<cellStyles count=\"1\">" +
            "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" />" +
            "</cellStyles>" +
            "<dxfs count=\"0\" />" +
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
                buffer.Index += Utf8Helper.GetBytes(Header, buffer.GetNextSpan());

                // Fonts
                ProcessFonts(styles, out var fontList, out var fontLookup);
                buffer.Index += Utf8Helper.GetBytes("<fonts count=\"", buffer.GetNextSpan());
                buffer.Index += Utf8Helper.GetBytes(fontList.Count, buffer.GetNextSpan());
                buffer.Index += Utf8Helper.GetBytes("\">", buffer.GetNextSpan());

                // The default font must be the first one (index 0)
                const string defaultFont = "<font><sz val=\"11\" /><name val=\"Calibri\" /></font>";
                await buffer.WriteAsciiStringAsync(defaultFont, stream, token).ConfigureAwait(false);

                var sb = new StringBuilder();
                for (var i = 1; i < fontList.Count; ++i)
                {
                    var font = fontList[i];
                    sb.Clear();
                    sb.Append("<font>");

                    if (font.Bold) sb.Append("<b/>");
                    if (font.Italic) sb.Append("<i/>");
                    if (font.Strikethrough) sb.Append("<strike/>");

                    sb.Append("<sz val=\"11\" /><name val=\"Calibri\" /></font>");
                    await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);
                }

                await buffer.WriteAsciiStringAsync(XmlPart1, stream, token).ConfigureAwait(false);

                var styleCount = styles.Count + 1;
                if (styleCount.GetNumberOfDigits() > buffer.GetRemainingBuffer())
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                // The default style must be the first one (index 0)
                const string defaultStyle = "\"><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" />";
                await buffer.WriteAsciiStringAsync(defaultStyle, stream, token).ConfigureAwait(false);

                sb.Clear();
                for (var i = 0; i < styles.Count; ++i)
                {
                    var fontIndex = fontLookup[i];
                    sb.Append("<xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"");
                    sb.Append(fontIndex);
                    sb.Append("\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" />");
                    await buffer.WriteAsciiStringAsync(sb.ToString(), stream, token).ConfigureAwait(false);
                }

                await buffer.WriteAsciiStringAsync(XmlPart2, stream, token).ConfigureAwait(false);
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }
        }

        private static void ProcessFonts(List<Style> styles, out List<Font> uniqueFonts, out List<int> fontLookup)
        {
            var defaultFont = new Font();
            var fontSet = new Dictionary<Font, int>(styles.Count) { { defaultFont, 0 } };
            uniqueFonts = new List<Font>(styles.Count) { defaultFont };
            fontLookup = new List<int>(styles.Count);

            for (var i = 0; i < styles.Count; ++i)
            {
                var font = styles[i].Font;

                // New unique font?
                if (!fontSet.TryGetValue(font, out var listIndex))
                {
                    listIndex = uniqueFonts.Count;
                    fontSet[font] = listIndex;
                    uniqueFonts.Add(font);
                }

                fontLookup.Add(listIndex);
            }
        }
    }
}
