using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal static class ContentTypesXml
{
    private const string Header =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />";

    private const string Styles = "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />";

    private const string SheetStartString = "<Override PartName=\"/";
    private const string SheetEndString = "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />";
    private const string Footer = "</Types>";

    private static readonly byte[] SheetStart = Utf8Helper.GetBytes(SheetStartString);
    private static readonly byte[] SheetEnd = Utf8Helper.GetBytes(SheetEndString);

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        List<WorksheetMetadata> worksheets,
        bool hasStylesXml,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("[Content_Types].xml", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            buffer.Advance(Utf8Helper.GetBytes(Header, buffer.GetSpan()));

            if (hasStylesXml)
                await buffer.WriteAsciiStringAsync(Styles, stream, token).ConfigureAwait(false);

            for (var i = 0; i < worksheets.Count; ++i)
            {
                var path = worksheets[i].Path;
                var sheetElementLength = GetSheetElementByteCount(path);

                if (sheetElementLength > buffer.FreeCapacity)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                buffer.Advance(GetSheetElementBytes(path, buffer.GetSpan()));
            }

            await buffer.WriteAsciiStringAsync(Footer, stream, token).ConfigureAwait(false);
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static int GetSheetElementByteCount(string path)
    {
        return SheetStart.Length
            + Utf8Helper.GetByteCount(path)
            + SheetEnd.Length;
    }

    private static int GetSheetElementBytes(string path, Span<byte> bytes)
    {
        var bytesWritten = SpanHelper.GetBytes(SheetStart, bytes);
        bytesWritten += Utf8Helper.GetBytes(path, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(SheetEnd, bytes.Slice(bytesWritten));
        return bytesWritten;
    }
}
