using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookXml
{
    public static async ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        CancellationToken token)
    {
        var stream = zipArchiveManager.OpenEntry("xl/workbook.xml");
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new WorkbookXml(worksheets, buffer);

            foreach (var success in writer)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetsStart => "<sheets>"u8;
    private static ReadOnlySpan<byte> SheetsEnd => "</sheets></workbook>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly SpreadsheetBuffer _buffer;
    private Element _next;
    private int _nextWorksheetIndex;

    private WorkbookXml(
        List<WorksheetMetadata> worksheets,
        SpreadsheetBuffer buffer)
    {
        _worksheets = worksheets;
        _buffer = buffer;
    }

    public readonly WorkbookXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.BookViews => TryWriteBookViews(),
            Element.SheetsStart => _buffer.TryWrite(SheetsStart),
            Element.Sheets => TryWriteWorksheets(),
            _ => _buffer.TryWrite(SheetsEnd)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteBookViews()
    {
        var firstVisibleWorksheetId = _worksheets.FindIndex(static x => x.Visibility == WorksheetVisibility.Visible);
        if (firstVisibleWorksheetId <= 0)
            return true;

        return _buffer.TryWrite(
            $"{"<bookViews><workbookView firstSheet=\""u8}" +
            $"{firstVisibleWorksheetId}" +
            $"{"\" activeTab=\""u8}" +
            $"{firstVisibleWorksheetId}" +
            $"{"\"/></bookViews>"u8}");
    }

    private bool TryWriteWorksheets()
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

            var ok = _buffer.TryWrite(
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
