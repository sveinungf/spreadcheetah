using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingRelsXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int drawingsFileIndex,
        List<WorksheetImage> images,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/_rels/drawing{drawingsFileIndex}.xml.rels");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new DrawingRelsXml(images);
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header => """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> ImageStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image"""u8;
    private static ReadOnlySpan<byte> Footer => """</Relationships>"""u8;

    private readonly List<WorksheetImage> _images;
    private Element _next;
    private int _nextImageIndex;

    private DrawingRelsXml(List<WorksheetImage> images)
    {
        _images = images;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Images && !Advance(TryWriteImages(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteImages(Span<byte> bytes, ref int bytesWritten)
    {
        var images = _images;

        for (; _nextImageIndex < images.Count; ++_nextImageIndex)
        {
            var index = _nextImageIndex;
            var imageId = images[index].Image.EmbeddedImageId;
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!ImageStart.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(imageId, span, ref written)) return false;
            if (!""".png" Id="rId"""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(index + 1, span, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Images,
        Footer,
        Done
    }
}
