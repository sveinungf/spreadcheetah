using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct VmlDrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        Dictionary<SingleCellRelativeReference, string> notes,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/drawings/vmlDrawing{notesFilesIndex}.vml");
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new VmlDrawingXml(notes.Keys.ToList());
        return writer.WriteAsync(entry, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"""u8;

    // TODO: Remove whitespace
    private static ReadOnlySpan<byte> ShapeStart => """
        <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
            <v:stroke joinstyle="miter"/>
            <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>
        <v:shape id="shape_0" type="#_x0000_t202" style="position:absolute;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">
            <v:fill color2="infoBackground [80]"/>
            <v:shadow color="none [81]" obscured="t"/>
            <x:ClientData ObjectType="Note">
                <x:MoveWithCells/>
                <x:SizeWithCells/>
                <x:AutoFill>False</x:AutoFill>
                <x:Anchor>
        """u8;

    private static ReadOnlySpan<byte> ShapeAfterAnchor => "</x:Anchor><x:Row>"u8;
    private static ReadOnlySpan<byte> ShapeAfterRow => "</x:Row><x:Column>"u8;
    private static ReadOnlySpan<byte> ShapeEnd => "</x:Column></x:ClientData></v:shape>"u8;
    private static ReadOnlySpan<byte> Footer => "</xml>"u8;

    private readonly List<SingleCellRelativeReference> _noteReferences;
    private Element _next;
    private int _nextIndex;

    private VmlDrawingXml(List<SingleCellRelativeReference> noteReferences)
    {
        _noteReferences = noteReferences;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Notes && !Advance(TryWriteNotes(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteNotes(Span<byte> bytes, ref int bytesWritten)
    {
        var references = _noteReferences;

        for (; _nextIndex < references.Count; ++_nextIndex)
        {
            var reference = references[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!ShapeStart.TryCopyTo(span, ref written)) return false;

            // From an example:
            // From (upper right coordinate of a rectangle)
            // [0] Left column (startingColumn + 1)
            // [1] Left column offset (always 15)
            // [2] Left row (startingRow)
            // [3] Left row offset (startingRow == 0 ? 2 : 10)
            // To (bottom right coordinate of a rectangle)
            // [4] Right column (startingColumn + 3)
            // [5] Right column offset (always 15)
            // [6] Right row (startingRow + 3)
            // [7] Right row offset (startingRow == 0 ? 16 : 4)

            // TODO: Excel generated this anchor for A1
            if (!"1,15,0,2,3,31,4,1"u8.TryCopyTo(span, ref written)) return false;

            if (!ShapeAfterAnchor.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(reference.Row - 1, span, ref written)) return false;
            if (!ShapeAfterRow.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(reference.Column - 1, span, ref written)) return false;
            if (!ShapeEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Notes,
        Footer,
        Done
    }
}
