using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal static class XmlWriterExtensions
{
    public static async ValueTask WriteAsync<TXmlWriter>(
        this TXmlWriter writer,
        ZipArchiveEntry entry,
        SpreadsheetBuffer buffer,
        CancellationToken token)
        where TXmlWriter : IXmlWriter
    {
        var stream = entry.Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var done = false;

            do
            {
                done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
                buffer.Advance(bytesWritten);
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            } while (!done);
        }
    }
}
