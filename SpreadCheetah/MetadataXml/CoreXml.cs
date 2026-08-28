using SpreadCheetah.Helpers;
using SpreadCheetah.Metadata;

namespace SpreadCheetah.MetadataXml;

internal static class CoreXml
{
    public static ValueTask WriteCoreXmlAsync(
        this ZipArchiveManager zipArchiveManager,
        DocumentProperties documentProperties,
        SpreadsheetBuffer buffer,
        CancellationToken token)
    {
        const string entryName = "docProps/core.xml";
        var writer = new CoreXmlWriter(documentProperties, buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct CoreXmlWriter(
    DocumentProperties documentProperties,
    SpreadsheetBuffer buffer)
    : IXmlWriter<CoreXmlWriter>
{
    private Element _next;
    private BufferWriteProgress _progress;

    public readonly CoreXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Title => TryWriteTitle(),
            Element.Subject => TryWriteSubject(),
            Element.Author => TryWriteAuthor(),
            _ => TryWriteFooter()
        };

        if (Current)
        {
            _progress = default;
            ++_next;
        }

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var header =
            """<?xml version="1.0" encoding="utf-8"?>"""u8 +
            """<cp:coreProperties """u8 +
            """xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" """u8 +
            """xmlns:dc="http://purl.org/dc/elements/1.1/" """u8 +
            """xmlns:dcterms="http://purl.org/dc/terms/" """u8 +
            """xmlns:dcmitype="http://purl.org/dc/dcmitype/" """u8 +
            """xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"""u8;

        return buffer.TryWrite(header);
    }

    private bool TryWriteTitle()
    {
        if (documentProperties.Title is not { } title)
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{"<dc:title>"u8}{title}{"</dc:title>"u8}");
    }

    private bool TryWriteSubject()
    {
        if (documentProperties.Subject is not { } subject)
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{"<dc:subject>"u8}{subject}{"</dc:subject>"u8}");
    }

    private bool TryWriteAuthor()
    {
        if (documentProperties.Author is not { } author)
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{"<dc:creator>"u8}{author}{"</dc:creator>"u8}");
    }

    private readonly bool TryWriteFooter()
    {
        return buffer.TryWrite(
            $"{"""<dcterms:created xsi:type="dcterms:W3CDTF">"""u8}" +
            $"{DateTime.UtcNow}" +
            $"{"</dcterms:created></cp:coreProperties>"u8}");
    }
}

file enum Element
{
    Header,
    Title,
    Subject,
    Author,
    Footer,
    Done
}