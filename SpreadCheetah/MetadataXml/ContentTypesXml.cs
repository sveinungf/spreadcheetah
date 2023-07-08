using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct ContentTypesXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("[Content_Types].xml", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new ContentTypesXml(worksheets, hasStylesXml);
            var done = false;

            do
            {
                done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
                buffer.Advance(bytesWritten);
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            } while (!done);
        }
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"""u8 +
        """<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />"""u8 +
        """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />"""u8;

    private static ReadOnlySpan<byte> Vml => "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"u8;
    private static ReadOnlySpan<byte> Styles => """<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />"""u8;
    private static ReadOnlySpan<byte> SheetStart => """<Override PartName="/"""u8;
    private static ReadOnlySpan<byte> SheetEnd => "\" "u8 + """ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />"""u8;

    // TODO: Before closing Types tag if there are notes. 1 per sheet with notes.
    private static ReadOnlySpan<byte> Comment => "<Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>"u8;
    private static ReadOnlySpan<byte> Footer => "</Types>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly bool _hasStylesXml;
    private Element _next;
    private int _nextWorksheetIndex;

    private ContentTypesXml(List<WorksheetMetadata> worksheets, bool hasStylesXml)
    {
        _worksheets = worksheets;
        _hasStylesXml = hasStylesXml;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Vml && !Advance(TryWriteVml(bytes, ref bytesWritten))) return false;
        if (_next == Element.Styles && !Advance(TryWriteStyles(bytes, ref bytesWritten))) return false;
        if (_next == Element.Worksheets && !Advance(TryWriteWorksheets(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteStyles(Span<byte> bytes, ref int bytesWritten)
        => !_hasStylesXml || Styles.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteVml(Span<byte> bytes, ref int bytesWritten)
    {
        var hasNotes = _worksheets.Exists(x => x.NotesFileIndex is not null);
        return !hasNotes || Vml.TryCopyTo(bytes, ref bytesWritten);
    }

    private bool TryWriteWorksheets(Span<byte> bytes, ref int bytesWritten)
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var path = worksheets[_nextWorksheetIndex].Path;
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!SheetStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(path, span, ref written)) return false;
            if (!SheetEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Vml,
        Styles,
        Worksheets,
        Footer,
        Done
    }
}
