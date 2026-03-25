using Polyfills;
using System.IO.Compression;

namespace SpreadCheetah.Test.Helpers;

internal static class ZipArchiveExtensions
{
    extension(ZipArchive zip)
    {
        public Task<Stream> GetDrawingXmlStreamAsync(CancellationToken token)
        {
            return zip.GetXmlStreamAsync("xl/drawings/drawing1.xml", token);
        }

        public Task<Stream> GetSheet1XmlStreamAsync(CancellationToken token)
        {
            return zip.GetXmlStreamAsync("xl/worksheets/sheet1.xml", token);
        }

        public Task<Stream> GetStylesXmlStreamAsync(CancellationToken token)
        {
            return zip.GetXmlStreamAsync("xl/styles.xml", token);
        }

        private Task<Stream> GetXmlStreamAsync(string entryName, CancellationToken token)
        {
            var entry = zip.GetEntry(entryName)
                ?? throw new InvalidOperationException($"{entryName} not found in the provided XLSX stream.");

            return entry.OpenAsync(token);
        }
    }
}
