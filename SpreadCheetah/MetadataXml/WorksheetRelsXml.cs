using SpreadCheetah.Helpers;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml;

[StructLayout(LayoutKind.Auto)]
internal struct WorksheetRelsXml : IXmlWriter<WorksheetRelsXml>
{
    public static ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        int worksheetIndex,
        FileCounter fileCounter,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/worksheets/_rels/sheet{worksheetIndex}.xml.rels");
        var writer = new WorksheetRelsXml(fileCounter, buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> Footer => "</Relationships>"u8;

    private readonly FileCounter _fileCounter;
    private readonly SpreadsheetBuffer _buffer;
    private Element _next;
    private int _nextIndex;

    private WorksheetRelsXml(FileCounter fileCounter, SpreadsheetBuffer buffer)
    {
        _fileCounter = fileCounter;
        _buffer = buffer;
    }

    public readonly WorksheetRelsXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.VmlDrawing => TryWriteVmlDrawing(),
            Element.Comments => TryWriteComments(),
            Element.Drawing => TryWriteDrawing(),
            Element.Tables => TryWriteTables(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteVmlDrawing()
    {
        if (_fileCounter.CurrentWorksheetNotesFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.VmlDrawing}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing"u8}" +
            $"{fileIndex}" +
            $"{""".vml"/>"""u8}");
    }

    private readonly bool TryWriteComments()
    {
        if (_fileCounter.CurrentWorksheetNotesFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.Comments}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments"u8}" +
            $"{fileIndex}" +
            $"{""".xml"/>"""u8}");
    }

    private readonly bool TryWriteDrawing()
    {
        if (_fileCounter.CurrentWorksheetDrawingsFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.Drawing}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing"u8}" +
            $"{fileIndex}" +
            $"{""".xml"/>"""u8}");
    }

    private bool TryWriteTables()
    {
        if (_fileCounter.CurrentWorksheetTableFileStartIndex is not { } fileStartIndex)
            return true;

        for (; _nextIndex < _fileCounter.CurrentWorksheetTableCount; ++_nextIndex)
        {
            var relationshipId = WorksheetRelationshipIds.TableStartId + _nextIndex;
            var tableIndex = fileStartIndex + _nextIndex;

            var success = _buffer.TryWrite(
                $"{"<Relationship Id=\"rId"u8}" +
                $"{relationshipId}" +
                $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table"u8}" +
                $"{tableIndex}" +
                $"{""".xml"/>"""u8}");

            if (!success)
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private enum Element
    {
        Header,
        VmlDrawing,
        Comments,
        Drawing,
        Tables,
        Footer,
        Done
    }
}
