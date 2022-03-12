using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

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

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<string> worksheetNames,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("xl/workbook.xml", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
            await using (stream.ConfigureAwait(false))
#endif
        {
            buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

            for (var i = 0; i < worksheetNames.Count; ++i)
            {
                var sheetId = i + 1;
                var name = WebUtility.HtmlEncode(worksheetNames[i]);
                var sheetElementLength = GetSheetElementByteCount(name, sheetId);

                if (sheetElementLength > buffer.FreeCapacity)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                buffer.Advance(GetSheetElementBytes(name, sheetId, buffer.GetSpan()));
            }

            await buffer.WriteAsciiStringAsync(Footer, stream, token).ConfigureAwait(false);
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
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
        bytesWritten += Utf8Helper.GetBytes(sheetId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(BetweenSheetIdAndRelationId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(SharedMetadata.RelationIdPrefix, bytes.Slice(bytesWritten));
        bytesWritten += Utf8Helper.GetBytes(sheetId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(SheetEnd, bytes.Slice(bytesWritten));
        return bytesWritten;
    }
}
