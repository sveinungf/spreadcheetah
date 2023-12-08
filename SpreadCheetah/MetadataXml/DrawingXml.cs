using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int drawingsFileIndex,
        List<WorksheetImage> images,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/drawing{drawingsFileIndex}.xml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new DrawingXml(images);
#pragma warning disable EPS06 // Hidden struct copy operation
        return writer.WriteAsync(entry, buffer, token);
#pragma warning restore EPS06 // Hidden struct copy operation
    }

    private static ReadOnlySpan<byte> Header => """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" """u8 +
        """xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> Footer => """</xdr:wsDr>"""u8;

    private readonly List<WorksheetImage> _images;
    private DrawingImageXml? _currentImageXmlWriter;
    private Element _next;
    private int _nextIndex;

    private DrawingXml(List<WorksheetImage> images)
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

        for (; _nextIndex < images.Count; ++_nextIndex)
        {
            var imageXmlWriter = _currentImageXmlWriter
                ?? new DrawingImageXml(images[_nextIndex], _nextIndex);

            var span = bytes.Slice(bytesWritten);
            var done = imageXmlWriter.TryWrite(span, out var written);
            bytesWritten += written;

            if (!done)
            {
                _currentImageXmlWriter = imageXmlWriter;
                return false;
            }

            _currentImageXmlWriter = null;
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
