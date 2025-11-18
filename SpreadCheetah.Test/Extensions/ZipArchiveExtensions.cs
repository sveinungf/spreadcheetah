using System.IO.Compression;

namespace SpreadCheetah.Test.Extensions;

internal static class ZipArchiveExtensions
{
    extension(ZipArchive)
    {
        public static Task<ZipArchive> CreateAsync(
            Stream stream,
            CancellationToken cancellationToken = default)
        {
            return ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: false, entryNameEncoding: null, cancellationToken);
        }
    }
}
