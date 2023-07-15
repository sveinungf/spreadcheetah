using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal struct CommentsXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        Dictionary<CellReference, string> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/comments{notesFilesIndex}.xml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new CommentsXml(notes.ToList());
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8 +
        """<authors><author></author></authors>"""u8 +
        """<commentList>"""u8;

    private static ReadOnlySpan<byte> CommentStart => "<comment ref=\""u8;
    private static ReadOnlySpan<byte> CommentAfterRef => "\""u8 + """authorId="0"><text><t>"""u8;
    private static ReadOnlySpan<byte> CommentEnd => "</t></text></comment>"u8;
    private static ReadOnlySpan<byte> Footer => "</commentList></comments>"u8;

    private readonly List<KeyValuePair<CellReference, string>> _notes;
    private Element _next;
    private int _nextIndex;

    private CommentsXml(List<KeyValuePair<CellReference, string>> notes)
    {
        _notes = notes;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Comments && !Advance(TryWriteComments(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteComments(Span<byte> bytes, ref int bytesWritten)
    {
        var notes = _notes;

        for (; _nextIndex < notes.Count; ++_nextIndex)
        {
            var (cellRef, note) = notes[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!CommentStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(cellRef.Reference, span, ref written)) return false;
            if (!CommentAfterRef.TryCopyTo(span, ref written)) return false;

            var xmlEncodedNote = WebUtility.HtmlEncode(note);
            if (!SpanHelper.TryWrite(xmlEncodedNote, span, ref written)) return false;
            if (!CommentEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
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
