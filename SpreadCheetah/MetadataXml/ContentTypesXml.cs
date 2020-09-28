using SpreadCheetah.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class ContentTypesXml
    {
        private const string Header =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />";

        private const string SheetStartString = "<Override PartName=\"/";
        private const string SheetEndString = "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />";
        private const string Footer = "</Types>";

        private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
        private static readonly byte[] SheetEnd = Utf8Helper.GetBytes(SheetEndString);

        public static async ValueTask WriteContentAsync(Stream stream, byte[] buffer, List<string> worksheetPaths, CancellationToken token)
        {
            var bytesWritten = Utf8Helper.GetBytes(Header, buffer.AsSpan());

            for (var i = 0; i < worksheetPaths.Count; ++i)
            {
                var path = worksheetPaths[i];
                var sheetElementLength = GetSheetElementByteCount(path);

                if (sheetElementLength > buffer.Length - bytesWritten)
                    await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

                bytesWritten += GetSheetElementBytes(path, buffer.AsSpan(bytesWritten));
            }

            if (Footer.Length > buffer.Length - bytesWritten)
                await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

            bytesWritten += Utf8Helper.GetBytes(Footer, buffer.AsSpan(bytesWritten));
            await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);
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
