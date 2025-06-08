using System.IO.Compression;

namespace SpreadCheetah.Test.Helpers;

internal static class ZipArchiveExtensions
{
    public static Stream GetDrawingXmlStream(this ZipArchive zip)
    {
        var entry = zip.GetEntry("xl/drawings/drawing1.xml")
            ?? throw new InvalidOperationException("Drawing XML not found in the provided XLSX stream.");

        return entry.Open();
    }
}
