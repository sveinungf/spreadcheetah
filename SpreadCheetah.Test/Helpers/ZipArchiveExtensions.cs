using Polyfills;
using System.IO.Compression;

namespace SpreadCheetah.Test.Helpers;

internal static class ZipArchiveExtensions
{
    public static Task<Stream> GetDrawingXmlStreamAsync(this ZipArchive zip, CancellationToken token)
    {
        var entry = zip.GetEntry("xl/drawings/drawing1.xml")
            ?? throw new InvalidOperationException("Drawing XML not found in the provided XLSX stream.");

        return entry.OpenAsync(token);
    }

    public static Task<Stream> GetStylesXmlStreamAsync(this ZipArchive zip, CancellationToken token)
    {
        var entry = zip.GetEntry("xl/styles.xml")
            ?? throw new InvalidOperationException("Styles XML not found in the provided XLSX stream.");

        return entry.OpenAsync(token);
    }
}
