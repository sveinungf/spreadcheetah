using SpreadCheetah.Helpers;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class ContentTypesXml
    {
        // TODO: Default for XML could be worksheet+xml instead of main+xml?
        private const string Header =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />";

        private const string Styles = "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />";

        private const string SheetStartString = "<Override PartName=\"/";
        private const string SheetEndString = "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />";
        private const string Footer = "</Types>";

        private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
        private static readonly byte[] SheetEnd = Utf8Helper.GetBytes(SheetEndString);

        public static async ValueTask WriteAsync(
            ZipArchive archive,
            CompressionLevel compressionLevel,
            SpreadsheetBuffer buffer,
            List<string> worksheetPaths,
            bool hasStylesXml,
            CancellationToken token)
        {
            var stream = archive.CreateEntry("[Content_Types].xml", compressionLevel).Open();
#if NETSTANDARD2_0
            using (stream)
#else
            await using (stream.ConfigureAwait(false))
#endif
            {
                buffer.Index = Utf8Helper.GetBytes(Header, buffer.GetNextSpan());

                if (hasStylesXml)
                {
                    if (Styles.Length > buffer.GetRemainingBuffer())
                        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                    buffer.Index += Utf8Helper.GetBytes(Styles, buffer.GetNextSpan());
                }

                for (var i = 0; i < worksheetPaths.Count; ++i)
                {
                    var path = worksheetPaths[i];
                    var sheetElementLength = GetSheetElementByteCount(path);

                    if (sheetElementLength > buffer.GetRemainingBuffer())
                        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                    buffer.Index += GetSheetElementBytes(path, buffer.GetNextSpan());
                }

                if (Footer.Length > buffer.GetRemainingBuffer())
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                buffer.Index += Utf8Helper.GetBytes(Footer, buffer.GetNextSpan());
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }
        }

        private static int GetSheetElementByteCount(string path)
        {
            return SheetStart.Length
                + Utf8Helper.GetByteCount(path)
                + SheetEnd.Length;
        }

        private static int GetSheetElementBytes(string path, Span<byte> bytes)
        {
            var bytesWritten = SpanHelper.GetBytes(SheetStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(path, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(SheetEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
