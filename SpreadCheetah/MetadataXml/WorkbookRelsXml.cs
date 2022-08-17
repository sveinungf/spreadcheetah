using SpreadCheetah.Helpers;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal static class WorkbookRelsXml
{
    private const string Header =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

    private const string StylesStartString = "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\" Id=\"";
    private const string SheetStartString = "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/";
    private const string BetweenPathAndRelationIdString = "\" Id=\"";
    private const string RelationEndString = "\" />";
    private const string Footer = "</Relationships>";

    private static readonly byte[] StylesStart = Utf8Helper.GetBytes(StylesStartString);
    private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
    private static readonly byte[] BetweenPathAndRelationId = Utf8Helper.GetBytes(BetweenPathAndRelationIdString);
    private static readonly byte[] RelationEnd = Utf8Helper.GetBytes(RelationEndString);

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<string> worksheetPaths,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

            for (var i = 0; i < worksheetPaths.Count; ++i)
            {
                var path = worksheetPaths[i];
                var sheetId = i + 1;
                var sheetElementLength = GetSheetElementByteCount(path, sheetId);

                if (sheetElementLength > buffer.FreeCapacity)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                buffer.Advance(GetSheetElementBytes(path, sheetId, buffer.GetSpan()));
            }

            var bufferNeeded = Footer.Length + MaxStylesXmlElementByteCount;
            if (bufferNeeded > buffer.FreeCapacity)
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            buffer.Advance(GetStylesXmlElementBytes(worksheetPaths.Count + 1, buffer.GetSpan()));
            buffer.Advance(Utf8Helper.GetBytes(Footer, buffer.GetSpan()));
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static int GetSheetElementByteCount(string path, int sheetId)
    {
        var sheetIdDigits = sheetId.GetNumberOfDigits();
        return SheetStart.Length
            + Utf8Helper.GetByteCount(path)
            + BetweenPathAndRelationId.Length
            + SharedMetadata.RelationIdPrefix.Length
            + sheetIdDigits
            + RelationEnd.Length;
    }

    private static int GetSheetElementBytes(string path, int sheetId, Span<byte> bytes)
    {
        var bytesWritten = SpanHelper.GetBytes(SheetStart, bytes);
        bytesWritten += Utf8Helper.GetBytes(path, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(BetweenPathAndRelationId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(SharedMetadata.RelationIdPrefix, bytes.Slice(bytesWritten));
        bytesWritten += Utf8Helper.GetBytes(sheetId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RelationEnd, bytes.Slice(bytesWritten));
        return bytesWritten;
    }

    private static readonly int MaxStylesXmlElementByteCount = StylesStartString.Length
        + SharedMetadata.RelationIdPrefix.Length
        + SpreadsheetConstants.SheetCountMaxDigits
        + RelationEnd.Length;

    private static int GetStylesXmlElementBytes(int relationId, Span<byte> bytes)
    {
        var bytesWritten = SpanHelper.GetBytes(StylesStart, bytes);
        bytesWritten += SpanHelper.GetBytes(SharedMetadata.RelationIdPrefix, bytes.Slice(bytesWritten));
        bytesWritten += Utf8Helper.GetBytes(relationId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RelationEnd, bytes.Slice(bytesWritten));
        return bytesWritten;
    }
}
