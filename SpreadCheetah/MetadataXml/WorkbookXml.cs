using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("xl/workbook.xml", compressionLevel);
        var writer = new WorkbookXml(worksheets);
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetsStart => "<sheets>"u8;
    private static ReadOnlySpan<byte> SheetsEnd => "</sheets></workbook>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private Element _next;
    private int _nextWorksheetIndex;

    private WorkbookXml(List<WorksheetMetadata> worksheets) => _worksheets = worksheets;

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.BookViews && !Advance(TryWriteBookViews(bytes, ref bytesWritten))) return false;
        if (_next == Element.SheetsStart && !Advance(SheetsStart.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Sheets && !Advance(TryWriteWorksheets(bytes, ref bytesWritten))) return false;
        if (_next == Element.SheetsEnd && !Advance(SheetsEnd.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    public bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteBookViews(Span<byte> bytes, ref int bytesWritten)
    {
        var firstVisibleWorksheetId = _worksheets.FindIndex(static x => x.Visibility == WorksheetVisibility.Visible);
        if (firstVisibleWorksheetId <= 0) return true;

        var written = 0;
        var span = bytes.Slice(bytesWritten);

        if (!"<bookViews><workbookView firstSheet=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(firstVisibleWorksheetId, span, ref written)) return false;
        if (!"\" activeTab=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(firstVisibleWorksheetId, span, ref written)) return false;
        if (!"\"/></bookViews>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteWorksheets(Span<byte> bytes, ref int bytesWritten)
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var index = _nextWorksheetIndex;
            var sheet = worksheets[index];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!"<sheet name=\""u8.TryCopyTo(span, ref written)) return false;

            var name = WebUtility.HtmlEncode(sheet.Name);
            if (!SpanHelper.TryWrite(name, span, ref written)) return false;
            if (!"\" sheetId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(index + 1, span, ref written)) return false;

            if (sheet.Visibility == WorksheetVisibility.Hidden && !"\" state=\"hidden"u8.TryCopyTo(span, ref written)) return false;
            if (sheet.Visibility == WorksheetVisibility.VeryHidden && !"\" state=\"veryHidden"u8.TryCopyTo(span, ref written)) return false;

            if (!"\" r:id=\"rId"u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(index + 1, span, ref written)) return false;
            if (!"\" />"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private enum Element
    {
        Header,
        BookViews,
        SheetsStart,
        Sheets,
        SheetsEnd,
        Done
    }
}
