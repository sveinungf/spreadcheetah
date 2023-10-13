using System.IO.Compression;

namespace SpreadCheetah.Images;

internal static class ZipArchiveExtensions
{
    public static async ValueTask<EmbeddedImage> CreateImageEntryAsync(
        this ZipArchive archive,
        Stream stream,
        ReadOnlyMemory<byte> header,
        ImageType type,
        CancellationToken token)
    {
        // TODO: Increment image file name
        // TODO: Correct file extension for JPG/PNG
        var entryName = type switch
        {
            ImageType.Png => "xl/media/image1.png",
            _ => throw new ArgumentOutOfRangeException(nameof(type)),
        };

        var entry = archive.CreateEntry(entryName);
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
                ImageType.Png => stream.CopyPngToAsync(entryStream, token),
                _ => new ValueTask<EmbeddedImage>(new EmbeddedImage(0, 0))
            };

            return await task.ConfigureAwait(false);
        }
    }

    private static async ValueTask<EmbeddedImage> CopyPngToAsync(this Stream source, Stream destination, CancellationToken token)
    {
        await source.CopyToAsync(destination, token).ConfigureAwait(false);
        // TODO: Parse from file
        return new EmbeddedImage(0, 0);
    }
}
