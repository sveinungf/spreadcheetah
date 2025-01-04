using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingRelsXml
{
    public static async ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        int drawingsFileIndex,
        List<WorksheetImage> images,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/_rels/drawing{drawingsFileIndex}.xml.rels");
        var stream = zipArchiveManager.OpenEntry(entryName);
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new DrawingRelsXml(images, buffer);

            foreach (var success in writer)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header => """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> ImageStart => """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image"""u8;
    private static ReadOnlySpan<byte> Footer => """</Relationships>"""u8;

    private readonly List<WorksheetImage> _images;
    private readonly SpreadsheetBuffer _buffer;
    private Element _next;
    private int _nextImageIndex;

    private DrawingRelsXml(List<WorksheetImage> images, SpreadsheetBuffer buffer)
    {
        _images = images;
        _buffer = buffer;
    }

    public readonly DrawingRelsXml GetEnumerator() => this;
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

        for (; _nextImageIndex < images.Count; ++_nextImageIndex)
        {
            var index = _nextImageIndex;
            var imageId = images[index].EmbeddedImage.Id;
            var success = _buffer.TryWrite(
                $"{ImageStart}" +
                $"{imageId}" +
                $"{""".png" Id="rId"""u8}" +
                $"{index + 1}" +
                $"{"\"/>"u8}");

            if (!success)
                return false;
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
