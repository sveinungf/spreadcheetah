using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.MetadataXml
{
    internal static class RelationshipsXml
    {
        private const string Content =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"rId1\" />" +
            "</Relationships>";

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
                buffer.Advance(Utf8Helper.GetBytes(Content, buffer.GetNextSpan()));
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }
        }
    }
}
