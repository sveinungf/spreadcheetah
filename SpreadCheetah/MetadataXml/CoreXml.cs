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
    private int _currentStringIndex;

    public readonly CoreXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.TitleStart => TryWriteTitleStart(),
            Element.Title => TryWriteString(documentProperties.Title),
            Element.TitleEnd => TryWriteTitleEnd(),
            Element.SubjectStart => TryWriteSubjectStart(),
            Element.Subject => TryWriteString(documentProperties.Subject),
            Element.SubjectEnd => TryWriteSubjectEnd(),
            Element.AuthorStart => TryWriteAuthorStart(),
            Element.Author => TryWriteString(documentProperties.Author),
            Element.AuthorEnd => TryWriteAuthorEnd(),
            _ => TryWriteFooter()
        };

        if (Current)
            ++_next;

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

    private readonly bool TryWriteTitleStart() => documentProperties.Title is null || buffer.TryWrite("<dc:title>"u8);
    private readonly bool TryWriteTitleEnd() => documentProperties.Title is null || buffer.TryWrite("</dc:title>"u8);
    private readonly bool TryWriteSubjectStart() => documentProperties.Subject is null || buffer.TryWrite("<dc:subject>"u8);
    private readonly bool TryWriteSubjectEnd() => documentProperties.Subject is null || buffer.TryWrite("</dc:subject>"u8);
    private readonly bool TryWriteAuthorStart() => documentProperties.Author is null || buffer.TryWrite("<dc:creator>"u8);
    private readonly bool TryWriteAuthorEnd() => documentProperties.Author is null || buffer.TryWrite("</dc:creator>"u8);

    private bool TryWriteString(string? value)
    {
        if (value is null)
            return true;

        if (buffer.WriteLongString(value, ref _currentStringIndex))
        {
            _currentStringIndex = 0;
            return true;
        }

        return false;
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
    TitleStart,
    Title,
    TitleEnd,
    SubjectStart,
    Subject,
    SubjectEnd,
    AuthorStart,
    Author,
    AuthorEnd,
    Footer,
    Done
}