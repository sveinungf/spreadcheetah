using SpreadCheetah.Helpers;
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class WorkbookRelsXml
    {
        private const string Header =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

        private const string SheetStartString = "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/";
        private const string BetweenPathAndRelationIdString = "\" Id=\"";
        private const string SheetEndString = "\" />";
        private const string Footer = "</Relationships>";

        private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
        private static readonly byte[] BetweenPathAndRelationId = Utf8Helper.GetBytes(BetweenPathAndRelationIdString);
        private static readonly byte[] SheetEnd = Utf8Helper.GetBytes(SheetEndString);

        public static async ValueTask WriteContentAsync(Stream stream, byte[] buffer, List<string> worksheetPaths, CancellationToken token)
        {
            var bytesWritten = Utf8Helper.GetBytes(Header, buffer.AsSpan());

            for (var i = 0; i < worksheetPaths.Count; ++i)
            {
                var path = worksheetPaths[i];
                var sheetId = i + 1;
                var sheetElementLength = GetSheetElementByteCount(path, sheetId);

                if (sheetElementLength > buffer.Length - bytesWritten)
                    await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

                bytesWritten += GetSheetElementBytes(path, sheetId, buffer.AsSpan(bytesWritten));
            }

            if (Footer.Length > buffer.Length - bytesWritten)
                await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

            bytesWritten += Utf8Helper.GetBytes(Footer, buffer.AsSpan(bytesWritten));
            await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);
        }

        private static int GetSheetElementByteCount(string path, int sheetId)
        {
            var sheetIdDigits = sheetId.GetNumberOfDigits();
            return SheetStart.Length
                + Utf8Helper.GetByteCount(path)
                + BetweenPathAndRelationId.Length
                + SharedMetadata.RelationIdPrefix.Length
                + sheetIdDigits
                + SheetEnd.Length;
        }

        private static int GetSheetElementBytes(string path, int sheetId, Span<byte> bytes)
        {
            var bytesWritten = SpanHelper.GetBytes(SheetStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(path, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(BetweenPathAndRelationId, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(SharedMetadata.RelationIdPrefix, bytes.Slice(bytesWritten));

            Utf8Formatter.TryFormat(sheetId, bytes.Slice(bytesWritten), out var sheetIdBytes);
            bytesWritten += sheetIdBytes;
            bytesWritten += SpanHelper.GetBytes(SheetEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
