using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct WorkbookRelsXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel);
        var writer = new WorkbookRelsXml(worksheets, hasStylesXml);
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/"""u8;
    private static ReadOnlySpan<byte> StylesStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId"""u8;
    private static ReadOnlySpan<byte> Footer => "</Relationships>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly bool _hasStylesXml;
    private Element _next;
    private int _nextWorksheetIndex;

    private WorkbookRelsXml(List<WorksheetMetadata> worksheets, bool hasStylesXml)
    {
        _worksheets = worksheets;
        _hasStylesXml = hasStylesXml;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Worksheets && !Advance(TryWriteWorksheets(bytes, ref bytesWritten))) return false;
        if (_next == Element.Styles && !Advance(TryWriteStyles(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteWorksheets(Span<byte> bytes, ref int bytesWritten)
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var index = _nextWorksheetIndex;
            var path = worksheets[index].Path;
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!SheetStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(path, span, ref written)) return false;
            if (!"\" Id=\"rId"u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(index + 1, span, ref written)) return false;
            if (!"\" />"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private readonly bool TryWriteStyles(Span<byte> bytes, ref int bytesWritten)
    {
        if (!_hasStylesXml) return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!StylesStart.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_worksheets.Count + 1, span, ref written)) return false;
        if (!"\" />"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
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
