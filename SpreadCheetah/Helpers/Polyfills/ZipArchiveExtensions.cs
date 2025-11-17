#if !NET10_0_OR_GREATER
using System.Text;

namespace System.IO.Compression;

internal static class ZipArchiveExtensions
{
    extension(ZipArchive zipArchive)
    {
        public static Task<ZipArchive> CreateAsync(
            Stream stream,
            ZipArchiveMode mode,
            bool leaveOpen,
            Encoding? entryNameEncoding,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
#pragma warning disable RS0030 // Do not use banned APIs
            var archive = new ZipArchive(stream, mode, leaveOpen, entryNameEncoding);
#pragma warning restore RS0030 // Do not use banned APIs
            return Task.FromResult(archive);
        }

        public ValueTask DisposeAsync()
        {
#pragma warning disable MA0042 // Do not use blocking calls in an async method
            zipArchive.Dispose();
#pragma warning restore MA0042 // Do not use blocking calls in an async method
            return new ValueTask();
        }
    }
}
#endif