using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.Buffers.Text;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal static class WorkbookXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
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

            while (!TryWrite(ref indexA, ref indexB, buffer, worksheets))
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static bool TryWrite(ref int indexA, ref int indexB, SpreadsheetBuffer buffer, List<WorksheetMetadata> worksheets)
    {
        if (indexA == 0)
        {
            if (!TryWriteSpan(Header, buffer)) return false;
            ++indexA;
        }

        if (indexA == 1)
        {
            var firstVisibleWorksheetId = worksheets.FindIndex(x => x.Visibility == WorksheetVisibility.Visible);
            if (firstVisibleWorksheetId > 0 && !TryGetWorkbookViewBytes(firstVisibleWorksheetId, buffer))
                return false;

            ++indexA;
        }

        if (indexA == 2)
        {
            if (!TryWriteSpan("<sheets>"u8, buffer)) return false;
            ++indexA;
        }

        if (indexA == 3)
        {
            for (; indexB < worksheets.Count; ++indexB)
            {
                var sheetId = indexB + 1;
                if (!TryGetSheetElementBytes(worksheets[indexB], sheetId, buffer)) return false;
            }

            ++indexA;
            indexB = 0;
        }

        if (indexA == 4)
        {
            if (!TryWriteSpan("</sheets></workbook>"u8, buffer)) return false;
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

    private static bool TryGetSheetElementBytes(WorksheetMetadata worksheet, int sheetId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = 0;

        if (!TryWriteSpan("<sheet name=\""u8, bytes, ref bytesWritten)) return false;

        var name = WebUtility.HtmlEncode(worksheet.Name);
        if (!TryWriteString(name, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan("\" sheetId=\""u8, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteInteger(sheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;

        if (worksheet.Visibility == WorksheetVisibility.Hidden
            && !TryWriteSpan("\" state=\"hidden"u8, bytes.Slice(bytesWritten), ref bytesWritten))
        {
            return false;
        }

        if (worksheet.Visibility == WorksheetVisibility.VeryHidden
            && !TryWriteSpan("\" state=\"veryHidden"u8, bytes.Slice(bytesWritten), ref bytesWritten))
        {
            return false;
        }

        if (!TryWriteSpan("\" r:id=\"rId"u8, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteInteger(sheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan("\" />"u8, bytes.Slice(bytesWritten), ref bytesWritten)) return false;

        buffer.Advance(bytesWritten);
        return true;
    }

    private static bool TryGetWorkbookViewBytes(int firstVisibleSheetId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = 0;

        if (!TryWriteSpan("<bookViews><workbookView firstSheet=\""u8, bytes, ref bytesWritten)) return false;
        if (!TryWriteInteger(firstVisibleSheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan("\" activeTab=\""u8, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteInteger(firstVisibleSheetId, bytes.Slice(bytesWritten), ref bytesWritten)) return false;
        if (!TryWriteSpan("\"/></bookViews>"u8, bytes.Slice(bytesWritten), ref bytesWritten)) return false;

        buffer.Advance(bytesWritten);
        return true;
    }
}
