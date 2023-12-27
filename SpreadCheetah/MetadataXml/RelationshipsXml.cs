using System.Diagnostics;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal static class RelationshipsXml
{
    private static ReadOnlySpan<byte> Content =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8 +
        """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="rId1" />"""u8 +
        """</Relationships>"""u8;

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("_rels/.rels", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var ok = buffer.TryWrite($"{Content}");
            Debug.Assert(ok);
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }
}
