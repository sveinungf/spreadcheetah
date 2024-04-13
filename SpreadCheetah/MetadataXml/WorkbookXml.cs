using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookXml : IBufferXmlWriter
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
        return writer.WriteToBufferAsync(entry, buffer, token);
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

    public bool TryWrite(SpreadsheetBuffer buffer)
    {
        if (_next == Element.Header && !Advance(buffer.TryWrite(Header))) return false;
        if (_next == Element.BookViews && !Advance(TryWriteBookViews(buffer))) return false;
        if (_next == Element.SheetsStart && !Advance(buffer.TryWrite(SheetsStart))) return false;
        if (_next == Element.Sheets && !Advance(TryWriteWorksheets(buffer))) return false;
        if (_next == Element.SheetsEnd && !Advance(buffer.TryWrite(SheetsEnd))) return false;

        return true;
    }

    public bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteBookViews(SpreadsheetBuffer buffer)
    {
        var firstVisibleWorksheetId = _worksheets.FindIndex(static x => x.Visibility == WorksheetVisibility.Visible);
        if (firstVisibleWorksheetId <= 0)
            return true;

        return buffer.TryWrite(
            $"{"<bookViews><workbookView firstSheet=\""u8}" +
            $"{firstVisibleWorksheetId}" +
            $"{"\" activeTab=\""u8}" +
            $"{firstVisibleWorksheetId}" +
            $"{"\"/></bookViews>"u8}");
    }

    private bool TryWriteWorksheets(SpreadsheetBuffer buffer)
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var index = _nextWorksheetIndex;
            var sheet = worksheets[index];

            var visibilitySpan = sheet.Visibility switch
            {
                WorksheetVisibility.Hidden => "\" state=\"hidden"u8,
                WorksheetVisibility.VeryHidden => "\" state=\"veryHidden"u8,
                _ => []
            };

            var ok = buffer.TryWrite(
                $"{"<sheet name=\""u8}{sheet.Name}" +
                $"{"\" sheetId=\""u8}{index + 1}" +
                $"{visibilitySpan}" +
                $"{"\" r:id=\"rId"u8}{index + 1}{"\" />"u8}");

            if (!ok)
                return false;
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
