using SpreadCheetah.Styling;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class StylesXml
    {
        public static async ValueTask WriteAsync(
            ZipArchive archive,
            string path,
            CompressionLevel compressionLevel,
            List<Style> styles,
            CancellationToken token)
        {
            var stream = archive.CreateEntry(path, compressionLevel).Open();
#if NETSTANDARD2_0
            using (stream)
#else
            await using (stream.ConfigureAwait(false))
#endif
            {
                var sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                sb.Append("<numFmts count=\"0\" />");

                // Fonts
                ProcessFonts(styles, out var fontList, out var fontLookup);
                sb.Append("<fonts count=\"");
                sb.Append(fontList.Count);
                sb.Append("\">");

                // The default font must be the first one (index 0)
                sb.Append("<font><sz val=\"11\" /><name val=\"Calibri\" /></font>");

                for (var i = 0; i < fontList.Count; ++i)
                {
                    var font = fontList[i];
                    sb.Append("<font>");

                    if (font.Bold)
                        sb.Append("<b />");

                    sb.Append("<sz val=\"11\" /><name val=\"Calibri\" /></font>");
                }

                sb.Append("</fonts>");
                sb.Append("<fills count=\"2\">");
                sb.Append("<fill><patternFill patternType=\"none\" /></fill>");
                sb.Append("<fill><patternFill patternType=\"gray125\" /></fill>");
                sb.Append("</fills>");
                sb.Append("<borders count=\"1\">");
                sb.Append("<border>");
                sb.Append("<left />");
                sb.Append("<right />");
                sb.Append("<top />");
                sb.Append("<bottom />");
                sb.Append("<diagonal />");
                sb.Append("</border>");
                sb.Append("</borders>");
                sb.Append("<cellStyleXfs count=\"1\">");
                sb.Append("<xf numFmtId=\"0\" fontId=\"0\" />");
                sb.Append("</cellStyleXfs>");
                sb.Append("<cellXfs count=\"");
                sb.Append(styles.Count + 1);
                sb.Append("\">");

                // The default style must be the first one (index 0)
                sb.Append("<xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" />");

                for (var i = 0; i < styles.Count; ++i)
                {
                    var fontIndex = fontLookup[i];

                    sb.Append("<xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"");
                    sb.Append(fontIndex);
                    sb.Append("\" fontId=\"1\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" />");
                }

                sb.Append("</cellXfs>");
                sb.Append("<cellStyles count=\"1\">");
                sb.Append("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" />");
                sb.Append("</cellStyles>");
                sb.Append("<dxfs count=\"0\" />");
                sb.Append("</styleSheet>");

                using var streamWriter = new StreamWriter(stream, new UTF8Encoding(false));
                await streamWriter.WriteAsync(sb.ToString()).ConfigureAwait(false);
            }
        }

        private static void ProcessFonts(List<Style> styles, out List<Font> uniqueFonts, out List<int> fontLookup)
        {
            // TODO: Font must implement Equals and Hashcode
            var fontSet = new Dictionary<Font, int>(styles.Count);
            uniqueFonts = new List<Font>(styles.Count);
            fontLookup = new List<int>(styles.Count);

            for (var i = 0; i <= styles.Count; ++i)
            {
                var font = styles[i].Font;

                // TODO: Handle the default font. Maybe just add it to the fontSet first?
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
