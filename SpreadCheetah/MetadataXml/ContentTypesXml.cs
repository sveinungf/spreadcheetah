using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.MetadataXml;

internal struct ContentTypesXmlWriter
{
    private static ReadOnlySpan<byte> Header =>
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8 +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"u8 +
        "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />"u8 +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />"u8;

    private static ReadOnlySpan<byte> Styles => "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />"u8;
    private static ReadOnlySpan<byte> SheetStart => "<Override PartName=\"/"u8;
    private static ReadOnlySpan<byte> SheetEnd => "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />"u8;
    private static ReadOnlySpan<byte> Footer => "</Types>"u8;

    private readonly List<WorksheetMetadata> _worksheets;
    private readonly bool _hasStylesXml;
    private Element _nextElement;
    private int _nextWorksheetIndex;

    public ContentTypesXmlWriter(List<WorksheetMetadata> worksheets, bool hasStylesXml)
    {
        _worksheets = worksheets;
        _hasStylesXml = hasStylesXml;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_nextElement == Element.Header && !Advance(TryWriteSpan(Header, bytes, ref bytesWritten))) return false;
        if (_nextElement == Element.Styles && !Advance(TryWriteSpan(Styles, bytes, ref bytesWritten))) return false;
        if (_nextElement == Element.Worksheets && !Advance(TryWriteWorksheets(bytes, ref bytesWritten))) return false;
        if (_nextElement == Element.Footer && !Advance(TryWriteSpan(Footer, bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
        {
            _nextElement = _nextElement switch
            {
                Element.Header => _hasStylesXml ? Element.Styles : Element.Worksheets,
                Element.Styles => Element.Worksheets,
                Element.Worksheets => Element.Footer,
                _ => Element.Done
            };
        }

        return success;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryWriteSpan(ReadOnlySpan<byte> source, Span<byte> bytes, ref int bytesWritten)
    {
        if (!source.TryCopyTo(bytes.Slice(bytesWritten))) return false;
        bytesWritten += source.Length;
        return true;
    }

    private bool TryWriteWorksheets(Span<byte> bytes, ref int bytesWritten)
    {
        var worksheets = _worksheets;

        for (; _nextWorksheetIndex < worksheets.Count; ++_nextWorksheetIndex)
        {
            var path = worksheets[_nextWorksheetIndex].Path;

            var span = bytes.Slice(bytesWritten);
            if (!SheetStart.TryCopyTo(span)) return false;

            span = span.Slice(SheetStart.Length);
            if (!Utf8Helper.TryGetBytes(path, span, out var pathBytes)) return false;

            span = span.Slice(pathBytes);
            if (!SheetEnd.TryCopyTo(span)) return false;

            bytesWritten += SheetStart.Length + pathBytes + SheetEnd.Length;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Styles,
        Worksheets,
        Footer,
        Done
    }
}

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
            var writer = new ContentTypesXmlWriter(worksheets, hasStylesXml);
            var done = false;

            do
            {
                done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
                buffer.Advance(bytesWritten);
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            } while (!done);
        }
    }

    public static async ValueTask WriteAsyncOld(
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
