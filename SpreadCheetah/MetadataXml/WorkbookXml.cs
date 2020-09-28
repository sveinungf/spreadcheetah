using SpreadCheetah.Helpers;
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class WorkbookXml
    {
        private const string Header =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets>";

        private const string SheetStartString = "<sheet name=\"";
        private const string BetweenNameAndSheetIdString = "\" sheetId=\"";
        private const string BetweenSheetIdAndRelationIdString = "\" r:id=\"";
        private const string SheetEndString = "\" />";
        private const string Footer = "</sheets></workbook>";

        private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
        private static readonly byte[] BetweenNameAndSheetId = Utf8Helper.GetBytes(BetweenNameAndSheetIdString);
        private static readonly byte[] BetweenSheetIdAndRelationId = Utf8Helper.GetBytes(BetweenSheetIdAndRelationIdString);
        private static readonly byte[] SheetEnd = Utf8Helper.GetBytes(SheetEndString);

        public static async ValueTask WriteContentAsync(Stream stream, byte[] buffer, List<string> worksheetNames, CancellationToken token)
        {
            var bytesWritten = Utf8Helper.GetBytes(Header, buffer.AsSpan());

            for (var i = 0; i < worksheetNames.Count; ++i)
            {
                var sheetId = i + 1;
                var name = WebUtility.HtmlEncode(worksheetNames[i]);
                var sheetElementLength = GetSheetElementByteCount(name, sheetId);

                if (sheetElementLength > buffer.Length - bytesWritten)
                    await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

                bytesWritten += GetSheetElementBytes(name, sheetId, buffer.AsSpan(bytesWritten));
            }

            if (Footer.Length > buffer.Length - bytesWritten)
                await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);

            bytesWritten += Utf8Helper.GetBytes(Footer, buffer.AsSpan(bytesWritten));
            await buffer.FlushToStreamAsync(stream, ref bytesWritten, token).ConfigureAwait(false);
        }

        private static int GetSheetElementByteCount(string name, int sheetId)
        {
            var sheetIdDigits = sheetId.GetNumberOfDigits();
            return SheetStart.Length
                + Utf8Helper.GetByteCount(name)
                + BetweenNameAndSheetId.Length
                + sheetIdDigits
                + BetweenSheetIdAndRelationId.Length
                + SharedMetadata.RelationIdPrefix.Length
                + sheetIdDigits
                + SheetEnd.Length;
        }

        private static int GetSheetElementBytes(string name, int sheetId, Span<byte> bytes)
        {
            var bytesWritten = SpanHelper.GetBytes(SheetStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(name, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(BetweenNameAndSheetId, bytes.Slice(bytesWritten));

            Utf8Formatter.TryFormat(sheetId, bytes.Slice(bytesWritten), out var sheetIdBytes);
            bytesWritten += sheetIdBytes;
            bytesWritten += SpanHelper.GetBytes(BetweenSheetIdAndRelationId, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(SharedMetadata.RelationIdPrefix, bytes.Slice(bytesWritten));

            Utf8Formatter.TryFormat(sheetId, bytes.Slice(bytesWritten), out sheetIdBytes);
            bytesWritten += sheetIdBytes;
            bytesWritten += SpanHelper.GetBytes(SheetEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
