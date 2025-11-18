using SpreadCheetah.Images;
using SpreadCheetah.MetadataXml;
using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;

namespace SpreadCheetah.Helpers;

internal sealed class ZipArchiveManager : IDisposable, IAsyncDisposable
{
    private readonly ZipArchive _zipArchive;
    private readonly CompressionLevel _compressionLevel;
    private bool _disposed;

    private ZipArchiveManager(
        ZipArchive zipArchive,
        SpreadCheetahCompressionLevel? compressionLevel)
    {
        _zipArchive = zipArchive;

        var level = compressionLevel ?? SpreadCheetahOptions.DefaultCompressionLevel;
        _compressionLevel = level is SpreadCheetahCompressionLevel.Optimal
            ? CompressionLevel.Optimal
            : CompressionLevel.Fastest;
    }

    public static async Task<ZipArchiveManager> CreateAsync(
        Stream stream,
        SpreadCheetahCompressionLevel? compressionLevel,
        CancellationToken token)
    {
        var zipArchive = await ZipArchive.CreateAsync(
            stream: stream,
            mode: ZipArchiveMode.Create,
            leaveOpen: true,
            entryNameEncoding: null,
            cancellationToken: token).ConfigureAwait(false);

        return new ZipArchiveManager(zipArchive, compressionLevel);
    }

    public Task<Stream> OpenEntryAsync(string entryName, CancellationToken token)
    {
        return _zipArchive.CreateEntry(entryName, _compressionLevel).OpenAsync(token);
    }

    public async ValueTask WriteAsync<T>(
        T xmlWriter,
        string entryName,
        SpreadsheetBuffer buffer,
        CancellationToken token)
        where T : struct, IXmlWriter<T>
    {
        var stream = await OpenEntryAsync(entryName, token).ConfigureAwait(false);
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            foreach (var success in xmlWriter)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
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
        var entryStream = await OpenEntryAsync(entryName, token).ConfigureAwait(false);
#if NETSTANDARD2_0
        using (entryStream)
#else
        await using (entryStream.ConfigureAwait(false))
#endif
        {
            await entryStream.WriteAsync(header, token).ConfigureAwait(false);

            Debug.Assert(type is ImageType.Png);
            return await CopyPngToAsync(stream, entryStream, header, embeddedImageId, spreadsheetGuid, token).ConfigureAwait(false);
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
        if (_disposed)
            return;

        _zipArchive.Dispose();
        _disposed = true;
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        await _zipArchive.DisposeAsync().ConfigureAwait(false);
        _disposed = true;
    }
}
