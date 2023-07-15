using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace SpreadCheetah.MetadataXml;

internal struct VmlDrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        int notesFilesIndex,
        Dictionary<CellReference, string> notes,
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

    private readonly List<CellReference> _noteReferences;
    private Element _next;
    private int _nextIndex;

    private VmlDrawingXml(List<CellReference> noteReferences)
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

            var (column, row) = ParseReference(reference.Reference);

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
            if (!SpanHelper.TryWrite(row - 1, span, ref written)) return false;
            if (!ShapeAfterRow.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(column - 1, span, ref written)) return false;
            if (!ShapeEnd.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    // TODO: GeneratedRegex on .NET 7
    private static readonly Regex Regex = new("^([A-Z]{1,3})([1-9][0-9]{0,6})$");

    /// <summary>"A1" returns (1,1)</summary>
    private static (int Column, int Row) ParseReference(string reference)
    {
        var match = Regex.Match(reference);
        if (!match.Success || match.Captures.Count < 2)
            throw new ArgumentException("Invalid reference.", nameof(reference));

        if (!TryParseColumnNumber(match.Captures[0], out var column) || !TryParseInteger(match.Captures[1], out var row))
            throw new ArgumentException("Invalid reference.", nameof(reference));

        return (column, row);
    }

    // TODO: Make optimized variant and move to SpreadsheetUtility
    private static bool TryParseColumnNumber(ReadOnlySpan<char> columnName, out int columnNumber)
    {
        columnNumber = 0;
        var pow = 1;
        for (int i = columnName.Length - 1; i >= 0; i--)
        {
            columnNumber += (columnName[i] - 'A' + 1) * pow;
            pow *= 26;
        }

        return true;
    }

    private static bool TryParseColumnNumber(Capture capture, out int columnNumber)
    {
#if NET6_0_OR_GREATER
        return TryParseColumnNumber(capture.ValueSpan, out columnNumber);
#else
        return TryParseColumnNumber(capture.Value.AsSpan(), out columnNumber);
#endif
    }

    private static bool TryParseInteger(Capture capture, out int result)
    {
#if NET6_0_OR_GREATER
        return int.TryParse(capture.ValueSpan, out result);
#else
        return int.TryParse(capture.Value, out result);
#endif
    }

    private enum Element
    {
        Header,
        Notes,
        Footer,
        Done
    }
}
