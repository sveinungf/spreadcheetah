using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct VmlDrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        Dictionary<SingleCellRelativeReference, string> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/vmlDrawing{notesFilesIndex}.vml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new VmlDrawingXml(notes.Keys.ToList());
#pragma warning disable EPS06 // Hidden struct copy operation
        return writer.WriteAsync(entry, buffer, token);
#pragma warning restore EPS06 // Hidden struct copy operation
    }

    private static ReadOnlySpan<byte> Header =>
        """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"""u8;

    private static ReadOnlySpan<byte> Footer => "</xml>"u8;

    private readonly List<SingleCellRelativeReference> _noteReferences;
    private VmlDrawingNoteXml? _currentNoteXmlWriter;
    private Element _next;
    private int _nextIndex;

    private VmlDrawingXml(List<SingleCellRelativeReference> noteReferences)
    {
        _noteReferences = noteReferences;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Notes && !Advance(TryWriteNotes(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteNotes(Span<byte> bytes, ref int bytesWritten)
    {
        var references = _noteReferences;

        for (; _nextIndex < references.Count; ++_nextIndex)
        {
            var span = bytes.Slice(bytesWritten);

            var noteXmlWriter = _currentNoteXmlWriter ?? new VmlDrawingNoteXml(references[_nextIndex]);
            if (!noteXmlWriter.TryWrite(span, out var written))
            {
                _currentNoteXmlWriter = noteXmlWriter;
                bytesWritten += written;
                return false;
            }

            _currentNoteXmlWriter = null;
            bytesWritten += written;
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
