using SpreadCheetah.Images;
using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;

namespace SpreadCheetah.Helpers;

internal sealed class ZipArchiveManager : IDisposable
{
    private readonly ZipArchive _zipArchive;
    private readonly CompressionLevel _compressionLevel;

    public ZipArchiveManager(
        Stream stream,
        SpreadCheetahCompressionLevel? compressionLevel)
    {
        _zipArchive = new ZipArchive(stream, ZipArchiveMode.Create, true);

        var level = compressionLevel ?? SpreadCheetahOptions.DefaultCompressionLevel;
        _compressionLevel = level is SpreadCheetahCompressionLevel.Optimal
            ? CompressionLevel.Optimal
            : CompressionLevel.Fastest;
    }

    public Stream OpenEntry(string entryName)
    {
        return _zipArchive.CreateEntry(entryName, _compressionLevel).Open();
    }

    public async ValueTask<EmbeddedImage> CreateImageEntryAsync(
        Stream stream,
        ReadOnlyMemory<byte> header,
        ImageType type,
        int embeddedImageId,
        Guid spreadsheetGuid,
        CancellationToken token)
    {
        var fileExtension = type switch
        {
            ImageType.Png => ".png",
            ImageType.Jpeg => ".jpeg",
            _ => throw new ArgumentOutOfRangeException(nameof(type)),
        };

        var entryName = StringHelper.Invariant($"xl/media/image{embeddedImageId}{fileExtension}");
        var entryStream = OpenEntry(entryName);
#if NETSTANDARD2_0
        using (entryStream)
#else
        await using (entryStream.ConfigureAwait(false))
#endif
        {
            await entryStream.WriteAsync(header, token).ConfigureAwait(false);

            var task = type switch
            {
                ImageType.Png => CopyPngToAsync(stream, entryStream, header, embeddedImageId, spreadsheetGuid, token),
                _ => new ValueTask<EmbeddedImage>(new EmbeddedImage(0, 0, embeddedImageId, spreadsheetGuid))
            };

            return await task.ConfigureAwait(false);
        }
    }

    private static async ValueTask<EmbeddedImage> CopyPngToAsync(Stream source, Stream destination, ReadOnlyMemory<byte> bytesReadSoFar,
        int embeddedImageId, Guid spreadsheetGuid, CancellationToken token)
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
        return new EmbeddedImage(width, height, embeddedImageId, spreadsheetGuid);
    }

    private static (int Width, int Height) ReadPngDimensions(ReadOnlySpan<byte> bytes)
    {
        var width = BinaryPrimitives.ReadInt32BigEndian(bytes);
        var height = BinaryPrimitives.ReadInt32BigEndian(bytes.Slice(4));

        width.EnsureValidImageDimension();
        height.EnsureValidImageDimension();

        return (width, height);
    }

    public void Dispose()
    {
        _zipArchive.Dispose();
    }
}
