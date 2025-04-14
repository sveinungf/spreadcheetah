using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml;

internal static class RelationshipsXml
{
    public static ValueTask WriteRelationshipsXmlAsync(
        this ZipArchiveManager zipArchiveManager,
        bool includeDocumentProperties,
        SpreadsheetBuffer buffer,
        CancellationToken token)
    {
        const string entryName = "_rels/.rels";
        var writer = new RelationshipsXmlWriter(includeDocumentProperties, buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct RelationshipsXmlWriter(
    bool includeDocumentProperties,
    SpreadsheetBuffer buffer)
    : IXmlWriter<RelationshipsXmlWriter>
{
    private Element _next;

    public readonly RelationshipsXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.DocProps => TryWriteDocProps(),
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
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""u8;

        return buffer.TryWrite(header);
    }

    private readonly bool TryWriteDocProps()
    {
        if (!includeDocumentProperties)
            return true;

        var content =
            """<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>"""u8 +
            """<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>"""u8;

        return buffer.TryWrite(content);
    }

    private readonly bool TryWriteFooter()
    {
        var footer =
            """<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"""u8 +
            """</Relationships>"""u8;

        return buffer.TryWrite(footer);
    }
}

file enum Element
{
    Header,
    DocProps,
    Footer,
    Done
}