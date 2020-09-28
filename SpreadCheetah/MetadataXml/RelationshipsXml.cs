using SpreadCheetah.Helpers;
using System;
using System.IO;
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

        public static ValueTask WriteContentAsync(Stream stream, byte[] buffer, CancellationToken token)
        {
            var bytesWritten = Utf8Helper.GetBytes(Content, buffer.AsSpan());
            return buffer.FlushToStreamAsync(stream, ref bytesWritten, token);
        }
    }
}
