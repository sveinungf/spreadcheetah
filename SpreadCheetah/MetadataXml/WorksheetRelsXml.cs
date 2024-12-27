using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml;

[StructLayout(LayoutKind.Auto)]
internal struct WorksheetRelsXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int worksheetIndex,
        FileCounter fileCounter,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/worksheets/_rels/sheet{worksheetIndex}.xml.rels");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var stream = entry.Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new WorksheetRelsXml(fileCounter, buffer);

            foreach (var success in writer)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

    // TODO: Tables
    private static ReadOnlySpan<byte> Footer => "</Relationships>"u8;

    private readonly FileCounter _fileCounter;
    private readonly SpreadsheetBuffer _buffer;
    private Element _next;

    private WorksheetRelsXml(FileCounter fileCounter, SpreadsheetBuffer buffer)
    {
        _fileCounter = fileCounter;
        _buffer = buffer;
    }

    public readonly WorksheetRelsXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.VmlDrawing => TryWriteVmlDrawing(),
            Element.Comments => TryWriteComments(),
            Element.Drawing => TryWriteDrawing(),
            _ => _buffer.TryWrite(Footer)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteVmlDrawing()
    {
        if (_fileCounter.CurrentWorksheetNotesFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.VmlDrawing}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing"u8}" +
            $"{fileIndex}" +
            $"{""".vml"/>"""u8}");
    }

    private readonly bool TryWriteComments()
    {
        if (_fileCounter.CurrentWorksheetNotesFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.Comments}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments"u8}" +
            $"{fileIndex}" +
            $"{""".xml"/>"""u8}");
    }

    private readonly bool TryWriteDrawing()
    {
        if (_fileCounter.CurrentWorksheetDrawingsFileIndex is not { } fileIndex)
            return true;

        return _buffer.TryWrite(
            $"{"<Relationship Id=\""u8}" +
            $"{WorksheetRelationshipIds.Drawing}" +
            $"{"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing"u8}" +
            $"{fileIndex}" +
            $"{""".xml"/>"""u8}");
    }

    private enum Element
    {
        Header,
        VmlDrawing,
        Comments,
        Drawing,
        Footer,
        Done
    }
}
