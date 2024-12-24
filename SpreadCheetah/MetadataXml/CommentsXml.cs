using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct CommentsXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/comments{notesFilesIndex}.xml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var stream = entry.Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new CommentsXml(notes, buffer);

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
        """<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8 +
        """<authors><author/></authors>"""u8 +
        """<commentList>"""u8;

    private static ReadOnlySpan<byte> CommentStart => "<comment ref=\""u8;
    private static ReadOnlySpan<byte> CommentAfterRef => "\" "u8 + """authorId="0"><text><r><t>"""u8;
    private static ReadOnlySpan<byte> CommentEnd => "</t></r></text></comment>"u8;
    private static ReadOnlySpan<byte> Footer => "</commentList></comments>"u8;

    private readonly ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> _notes;
    private readonly SpreadsheetBuffer _buffer;
    private string? _currentXmlEncodedNote;
    private int _currentXmlEncodedNoteIndex;
    private Element _next;
    private int _nextIndex;

    private CommentsXml(
        ReadOnlyMemory<KeyValuePair<SingleCellRelativeReference, string>> notes,
        SpreadsheetBuffer buffer)
    {
        _notes = notes;
        _buffer = buffer;
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
            var span = _buffer.GetSpan();
            var written = 0;

            if (_currentXmlEncodedNote is null)
            {
                if (!CommentStart.TryCopyTo(span, ref written)) return false;
                if (!SpanHelper.TryWriteCellReference(cellRef.Column, cellRef.Row, span, ref written)) return false;
                if (!CommentAfterRef.TryCopyTo(span, ref written)) return false;

                _currentXmlEncodedNote = XmlUtility.XmlEncode(note);
                _currentXmlEncodedNoteIndex = 0;
            }

            if (!SpanHelper.TryWriteLongString(_currentXmlEncodedNote, ref _currentXmlEncodedNoteIndex, span, ref written))
            {
                _buffer.Advance(written);
                return false;
            }

            _buffer.Advance(written);

            if (!_buffer.TryWrite(CommentEnd))
                return false;

            _currentXmlEncodedNote = null;
            _currentXmlEncodedNoteIndex = 0;
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
