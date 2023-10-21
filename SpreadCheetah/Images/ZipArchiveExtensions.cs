using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;

namespace SpreadCheetah.Images;

internal static class ZipArchiveExtensions
{
    public static async ValueTask<EmbeddedImage> CreateImageEntryAsync(
        this ZipArchive archive,
        Stream stream,
        ReadOnlyMemory<byte> header,
        ImageType type,
        int embeddedImageId,
        CancellationToken token)
    {
        // TODO: Increment image file name. It should basically be sequential numbering without taking filetype into account.
        // TODO: Correct file extension for JPG/PNG
        var entryName = type switch
        {
            ImageType.Png => "xl/media/image1.png",
            _ => throw new ArgumentOutOfRangeException(nameof(type)),
        };

        var entry = archive.CreateEntry(entryName); // TODO: Compression level. Ban CreateEntry without it?
        var entryStream = entry.Open();
#if NETSTANDARD2_0
        using (entryStream)
#else
        await using (entryStream.ConfigureAwait(false))
#endif
        {
            await entryStream.WriteAsync(header, token).ConfigureAwait(false);

            var task = type switch
            {
                ImageType.Png => stream.CopyPngToAsync(entryStream, header, embeddedImageId, token),
                _ => new ValueTask<EmbeddedImage>(new EmbeddedImage(0, 0, embeddedImageId))
            };

            return await task.ConfigureAwait(false);
        }
    }

    private static async ValueTask<EmbeddedImage> CopyPngToAsync(this Stream source, Stream destination, ReadOnlyMemory<byte> bytesReadSoFar, int embeddedImageId, CancellationToken token)
    {
        // A valid PNG file should start with these bytes:
        // 8 bytes: File signature
        // 4 bytes: Chunk length
        // 4 bytes: Chunk type (IHDR)
        // 4 bytes: Image width
        // 4 bytes: Image height
        const int bytesBeforeDimensionStart = 16;
        const int bytesRequiredToReadDimensions = 24;
        Debug.Assert(bytesReadSoFar.Length >= bytesRequiredToReadDimensions);

        var (width, height) = ReadPngDimensions(bytesReadSoFar.Span.Slice(bytesBeforeDimensionStart));
        await source.CopyToAsync(destination, token).ConfigureAwait(false);
        return new EmbeddedImage(width, height, embeddedImageId);
    }

    private static (int Width, int Height) ReadPngDimensions(ReadOnlySpan<byte> bytes)
    {
        var width = BinaryPrimitives.ReadInt32BigEndian(bytes);
        var height = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(4));
        return (width, height);
    }
}
