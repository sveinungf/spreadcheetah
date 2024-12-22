using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookRelsXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel);
        var stream = entry.Open();
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
            var span = _buffer.GetSpan();
            var written = 0;

            if (!SheetStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(path, span, ref written)) return false;
            if (!"\" Id=\"rId"u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(index + 1, span, ref written)) return false;
            if (!"\" />"u8.TryCopyTo(span, ref written)) return false;

            _buffer.Advance(written);
        }

        return true;
    }

    private readonly bool TryWriteStyles()
    {
        if (!_hasStylesXml) return true;

        var span = _buffer.GetSpan();
        var written = 0;

        if (!StylesStart.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_worksheets.Count + 1, span, ref written)) return false;
        if (!"\" />"u8.TryCopyTo(span, ref written)) return false;

        _buffer.Advance(written);
        return true;
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
