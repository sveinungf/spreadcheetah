using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml;

internal struct CommentsXml : IXmlWriter<CommentsXml>
{
    public static ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        InlineXmlTags inlineXmlTags,
        int notesFilesIndex,
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/comments{notesFilesIndex}.xml");
        var writer = new CommentsXml(notes, buffer, inlineXmlTags);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8 +
        """<authors><author/></authors>"""u8 +
        """<commentList>"""u8;

    private static ReadOnlySpan<byte> CommentStart => "<comment ref=\""u8;
    private static ReadOnlySpan<byte> CommentEnd => "</t></r></text></comment>"u8;
    private static ReadOnlySpan<byte> Footer => "</commentList></comments>"u8;

    private readonly ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> _notes;
    private readonly SpreadsheetBuffer _buffer;
    private readonly InlineXmlTags _inlineXmlTags;
    private Element _next;
    private int _nextIndex;
    private BufferWriteProgress _progress;

    private CommentsXml(
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        SpreadsheetBuffer buffer,
        InlineXmlTags inlineXmlTags)
    {
        _notes = notes;
        _buffer = buffer;
        _inlineXmlTags = inlineXmlTags;
    }

    public readonly CommentsXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.Comments => TryWriteComments(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private bool TryWriteComments()
    {
        var notes = _notes.Span;

        for (; _nextIndex < notes.Length; ++_nextIndex)
        {
            var (cellRef, note) = notes[_nextIndex];
            var reference = new SimpleSingleCellReference(cellRef.Column, cellRef.Row);

            if (!_buffer.TryWrite(
                    _progress, out _progress,
                    $"{CommentStart}{reference}{_inlineXmlTags.CommentAfterRef}{note}{CommentEnd}"))
            {
                return false;
            }

            _progress = default;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Comments,
        Footer,
        Done
    }
}
