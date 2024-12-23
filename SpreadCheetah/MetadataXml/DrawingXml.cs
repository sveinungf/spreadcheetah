using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int drawingsFileIndex,
        List<WorksheetImage> images,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/drawing{drawingsFileIndex}.xml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var stream = entry.Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new DrawingXml(images, buffer);

            foreach (var success in writer)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header => """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" """u8 +
        """xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> Footer => """</xdr:wsDr>"""u8;

    private readonly List<WorksheetImage> _images;
    private readonly SpreadsheetBuffer _buffer;
    private DrawingImageXml? _currentImageXmlWriter;
    private Element _next;
    private int _nextIndex;

    private DrawingXml(List<WorksheetImage> images, SpreadsheetBuffer buffer)
    {
        _images = images;
        _buffer = buffer;
    }

    public readonly DrawingXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.Images => TryWriteImages(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private bool TryWriteImages()
    {
        var images = _images;

        for (; _nextIndex < images.Count; ++_nextIndex)
        {
            var writer = _currentImageXmlWriter
                ?? new DrawingImageXml(images[_nextIndex], _nextIndex, _buffer);

            if (!writer.TryWrite())
            {
                _currentImageXmlWriter = writer;
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
