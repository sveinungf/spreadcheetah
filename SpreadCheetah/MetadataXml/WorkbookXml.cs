using SpreadCheetah.Helpers;
using System.Buffers.Text;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal static class WorkbookXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8 +
        """<sheets>"""u8;

    private static ReadOnlySpan<byte> SheetStart => "<sheet name=\""u8;
    private static ReadOnlySpan<byte> BetweenNameAndSheetId => "\" sheetId=\""u8;
    private static ReadOnlySpan<byte> BetweenSheetIdAndRelationId => "\" r:id=\"rId"u8;
    private static ReadOnlySpan<byte> SheetEnd => "\" />"u8;
    private static ReadOnlySpan<byte> Footer => "</sheets></workbook>"u8;

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
            var indexA = 0;
            var indexB = 0;

            while (!TryWrite(ref indexA, ref indexB, buffer, worksheetNames))
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static bool TryWrite(ref int indexA, ref int indexB, SpreadsheetBuffer buffer, List<string> worksheetNames)
    {
        if (indexA == 0)
        {
            if (!TryWriteSpan(Header, buffer)) return false;
            ++indexA;
        }

        if (indexA == 1)
        {
            for (; indexB < worksheetNames.Count; ++indexB)
            {
                var sheetId = indexB + 1;
                var name = WebUtility.HtmlEncode(worksheetNames[indexB]);
                if (!TryGetSheetElementBytes(name, sheetId, buffer)) return false;
            }

            ++indexA;
            indexB = 0;
        }

        if (indexA == 2)
        {
            if (!TryWriteSpan(Footer, buffer)) return false;
            ++indexA;
        }

        return true;
    }

    private static bool TryWriteSpan(ReadOnlySpan<byte> span, SpreadsheetBuffer buffer)
    {
        if (!span.TryCopyTo(buffer.GetSpan())) return false;
        buffer.Advance(span.Length);
        return true;
    }

    private static bool TryWriteSpan(ReadOnlySpan<byte> span, Span<byte> bytes, ref int bytesWritten)
    {
        if (!span.TryCopyTo(bytes)) return false;
        bytesWritten += span.Length;
        return true;
    }

    private static bool TryWriteString(string value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Helper.TryGetBytes(value, bytes, out var valueLength)) return false;
        bytesWritten += valueLength;
        return true;
    }

    private static bool TryWriteInteger(int value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Formatter.TryFormat(value, bytes, out var valueLength)) return false;
        bytesWritten += valueLength;
        return true;
    }

    private static bool TryGetSheetElementBytes(string name, int sheetId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = 0;

        if (!TryWriteSpan(SheetStart, bytes, ref bytesWritten)) return false;
        if (!TryWriteString(name, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan(BetweenNameAndSheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteInteger(sheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan(BetweenSheetIdAndRelationId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteInteger(sheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan(SheetEnd, bytes.Slice(bytesWritten), ref bytesWritten)) return false;

        buffer.Advance(bytesWritten);
        return true;
    }
}
