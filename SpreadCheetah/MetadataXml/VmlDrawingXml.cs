using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml;

[StructLayout(LayoutKind.Auto)]
internal struct VmlDrawingXml : IXmlWriter<VmlDrawingXml>
{
    public static ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/vmlDrawing{notesFilesIndex}.vml");
        var writer = new VmlDrawingXml(notes, buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"""u8;

    private static ReadOnlySpan<byte> Footer => "</xml>"u8;

    private readonly ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> _notes;
    private readonly SpreadsheetBuffer _buffer;
    private VmlDrawingNoteXml? _currentNoteXmlWriter;
    private Element _next;
    private int _nextIndex;

    private VmlDrawingXml(
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        SpreadsheetBuffer buffer)
    {
        _notes = notes;
        _buffer = buffer;
    }

    public readonly VmlDrawingXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.Notes => TryWriteNotes(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private bool TryWriteNotes()
    {
        var notes = _notes.Span;

        for (; _nextIndex < notes.Length; ++_nextIndex)
        {
            var writer = _currentNoteXmlWriter
                ?? new VmlDrawingNoteXml(notes[_nextIndex].Key, _buffer);

            if (!writer.TryWrite())
            {
                _currentNoteXmlWriter = writer;
                return false;
            }

            _currentNoteXmlWriter = null;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Notes,
        Footer,
        Done
    }
}
