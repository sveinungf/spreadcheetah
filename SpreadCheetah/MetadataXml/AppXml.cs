using SpreadCheetah.Helpers;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml;

internal static class AppXml
{
    public static ValueTask WriteAppXmlAsync(
        this ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        CancellationToken token)
    {
        const string entryName = "docProps/app.xml";
        var writer = new AppXmlWriter(buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct AppXmlWriter(SpreadsheetBuffer buffer)
    : IXmlWriter<AppXmlWriter>
{
    public readonly AppXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = TryWriteContent();
        return false;
    }

    private readonly bool TryWriteContent()
    {
        var content =
            """<?xml version="1.0" encoding="utf-8"?>"""u8 +
            """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" """u8 +
            """xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"""u8 +
            """<Application>SpreadCheetah</Application>"""u8 +
            """</Properties>"""u8;

        Debug.Assert(content.Length <= SpreadCheetahOptions.MinimumBufferSize);

        return buffer.TryWrite(content);
    }
}