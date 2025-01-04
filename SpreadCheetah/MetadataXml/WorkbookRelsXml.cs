using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookRelsXml
{
    public static async ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        CancellationToken token)
    {
        var stream = zipArchiveManager.OpenEntry("xl/_rels/workbook.xml.rels");
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new WorkbookRelsXml(worksheets, hasStylesXml, buffer);

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
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/"""u8;
    private static ReadOnlySpan<byte> StylesStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId"""u8;
    private static ReadOnlySpan<byte> Footer => "</Relationships>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly SpreadsheetBuffer _buffer;
    private readonly bool _hasStylesXml;
    private Element _next;
    private int _nextWorksheetIndex;

    private WorkbookRelsXml(
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        SpreadsheetBuffer buffer)
    {
        _worksheets = worksheets;
        _hasStylesXml = hasStylesXml;
        _buffer = buffer;
    }

    public readonly WorkbookRelsXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.Worksheets => TryWriteWorksheets(),
            Element.Styles => TryWriteStyles(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private bool TryWriteWorksheets()
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var index = _nextWorksheetIndex;
            var path = worksheets[index].Path;
            var success = _buffer.TryWrite(
                $"{SheetStart}" +
                $"{path}" +
                $"{"\" Id=\"rId"u8}" +
                $"{index + 1}" +
                $"{"\" />"u8}");

            if (!success)
                return false;
        }

        return true;
    }

    private readonly bool TryWriteStyles()
    {
        return !_hasStylesXml
            || _buffer.TryWrite($"{StylesStart}{_worksheets.Count + 1}{"\" />"u8}");
    }

    private enum Element
    {
        Header,
        Worksheets,
        Styles,
        Footer,
        Done
    }
}
