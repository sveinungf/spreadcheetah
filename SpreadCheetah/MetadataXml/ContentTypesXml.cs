using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct ContentTypesXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        FileCounter? fileCounter,
        bool hasStylesXml,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("[Content_Types].xml", compressionLevel);
        var writer = new ContentTypesXml(worksheets, fileCounter, hasStylesXml);
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"""u8 +
        """<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />"""u8 +
        """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />"""u8;

    // TODO: If there are tables, add for each table file:
    // <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>

    private static ReadOnlySpan<byte> Jpeg => """<Default Extension="jpeg" ContentType="image/jpeg"/>"""u8;
    private static ReadOnlySpan<byte> Png => """<Default Extension="png" ContentType="image/png"/>"""u8;
    private static ReadOnlySpan<byte> Vml => "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"u8;
    private static ReadOnlySpan<byte> Styles => """<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />"""u8;
    private static ReadOnlySpan<byte> DrawingStart => """<Override PartName="/xl/drawings/drawing"""u8;
    private static ReadOnlySpan<byte> DrawingEnd => """.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"""u8;
    private static ReadOnlySpan<byte> SheetStart => """<Override PartName="/"""u8;
    private static ReadOnlySpan<byte> SheetEnd => "\" "u8 + """ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />"""u8;
    private static ReadOnlySpan<byte> CommentStart => """<Override PartName="/xl/comments"""u8;
    private static ReadOnlySpan<byte> CommentEnd => """.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"""u8;
    private static ReadOnlySpan<byte> Footer => "</Types>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly FileCounter? _fileCounter;
    private readonly bool _hasStylesXml;
    private Element _next;
    private int _nextIndex;

    private ContentTypesXml(List<WorksheetMetadata> worksheets, FileCounter? fileCounter, bool hasStylesXml)
    {
        _worksheets = worksheets;
        _fileCounter = fileCounter;
        _hasStylesXml = hasStylesXml;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.ImageTypes && !Advance(TryWriteImageTypes(bytes, ref bytesWritten))) return false;
        if (_next == Element.Vml && !Advance(TryWriteVml(bytes, ref bytesWritten))) return false;
        if (_next == Element.Styles && !Advance(TryWriteStyles(bytes, ref bytesWritten))) return false;
        if (_next == Element.Drawings && !Advance(TryWriteDrawings(bytes, ref bytesWritten))) return false;
        if (_next == Element.Worksheets && !Advance(TryWriteWorksheets(bytes, ref bytesWritten))) return false;
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

    private readonly bool TryWriteImageTypes(Span<byte> bytes, ref int bytesWritten)
    {
        if (_fileCounter is not { } counter)
            return true;

        if (counter.EmbeddedImageTypes.HasFlag(EmbeddedImageTypes.Png) && !Png.TryCopyTo(bytes, ref bytesWritten))
            return false;

        if (counter.EmbeddedImageTypes.HasFlag(EmbeddedImageTypes.Jpeg) && !Jpeg.TryCopyTo(bytes, ref bytesWritten))
            return false;

        return true;
    }

    private readonly bool TryWriteStyles(Span<byte> bytes, ref int bytesWritten)
        => !_hasStylesXml || Styles.TryCopyTo(bytes, ref bytesWritten);

    private bool TryWriteDrawings(Span<byte> bytes, ref int bytesWritten)
    {
        if (_fileCounter is not { } counter)
            return true;

        for (; _nextIndex < counter.WorksheetsWithImages; ++_nextIndex)
        {
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!DrawingStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(_nextIndex + 1, span, ref written)) return false;
            if (!DrawingEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteVml(Span<byte> bytes, ref int bytesWritten)
    {
        var hasNotes = _fileCounter is { WorksheetsWithNotes: > 0 };
        return !hasNotes || Vml.TryCopyTo(bytes, ref bytesWritten);
    }

    private bool TryWriteWorksheets(Span<byte> bytes, ref int bytesWritten)
    {
        var worksheets = _worksheets;

        for (; _nextIndex < worksheets.Count; ++_nextIndex)
        {
            var worksheet = worksheets[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!SheetStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(worksheet.Path, span, ref written)) return false;
            if (!SheetEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        _nextIndex = 0;
        return true;
    }

    private bool TryWriteComments(Span<byte> bytes, ref int bytesWritten)
    {
        if (_fileCounter is not { } counter)
            return true;

        for (; _nextIndex < counter.WorksheetsWithNotes; ++_nextIndex)
        {
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!CommentStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(_nextIndex + 1, span, ref written)) return false;
            if (!CommentEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        _nextIndex = 0;
        return true;
    }

    private enum Element
    {
        Header,
        ImageTypes,
        Vml,
        Styles,
        Drawings,
        Worksheets,
        Comments,
        Footer,
        Done
    }
}
