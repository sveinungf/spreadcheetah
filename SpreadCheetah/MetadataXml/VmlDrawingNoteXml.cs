using SpreadCheetah.CellReferences;
using System.Buffers;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml;

[StructLayout(LayoutKind.Auto)]
internal struct VmlDrawingNoteXml(
    SingleCellRelativeReference reference,
    SpreadsheetBuffer buffer)
{
    private static ReadOnlySpan<byte> ShapeStart =>
        """<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">"""u8
        + """<v:stroke joinstyle="miter"/>"""u8
        + """<v:path gradientshapeok="t" o:connecttype="rect"/>"""u8
        + """</v:shapetype>"""u8
        + """<v:shape id="shape_0" type="#_x0000_t202" style="position:absolute;margin-left:57pt;margin-top:"""u8;

    private static ReadOnlySpan<byte> ShapeAfterMarginTop =>
        """pt;width:100.8pt;height:60.6pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">"""u8
        + """<v:fill color2="infoBackground [80]"/>"""u8
        + """<v:shadow color="none [81]" obscured="t"/>"""u8
        + """<v:textbox/>"""u8
        + """<x:ClientData ObjectType="Note">"""u8
        + """<x:MoveWithCells/>"""u8
        + """<x:SizeWithCells/>"""u8
        + """<x:AutoFill>False</x:AutoFill>"""u8
        + """<x:Anchor>"""u8;

    private static ReadOnlySpan<byte> ShapeAfterAnchor => "</x:Anchor><x:Row>"u8;
    private static ReadOnlySpan<byte> ShapeAfterRow => "</x:Row><x:Column>"u8;
    private static ReadOnlySpan<byte> ShapeEnd => "</x:Column></x:ClientData></v:shape>"u8;

    private Element _next;

    public bool TryWrite()
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.ShapeStart => buffer.TryWrite(ShapeStart),
            Element.MarginTop => TryWriteMarginTop(),
            Element.ShapeAfterMarginTop => buffer.TryWrite(ShapeAfterMarginTop),
            Element.Anchor => TryWriteAnchor(),
            _ => TryWriteShapeEnd()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteMarginTop()
    {
        var row = reference.Row;
        var marginTop = row == 1 ? 1.2 : row * 14.4 - 20.4;
        var tuple = (marginTop, new StandardFormat('F', 1));
        return buffer.TryWrite($"{tuple}");
    }

    /// <summary>
    /// From
    /// [0] Left column
    /// [1] Left column offset
    /// [2] Left row
    /// [3] Left row offset
    /// To
    /// [4] Right column
    /// [5] Right column offset
    /// [6] Right row
    /// [7] Right row offset
    /// </summary>
    private readonly bool TryWriteAnchor()
    {
        var col = reference.Column;
        var row = reference.Row;

        if (row <= 1)
        {
            return buffer.TryWrite(
                $"{col}" +
                $"{",12,0,1,"u8}" +
                $"{(ushort)(col + 2)}" +
                $"{",18,4,5"u8}");
        }

        return buffer.TryWrite(
            $"{col}" +
            $"{",12,"u8}" +
            $"{row - 2}" +
            $"{",11,"u8}" +
            $"{(ushort)(col + 2)}" +
            $"{",18,"u8}" +
            $"{row + 2}" +
            $"{",15"u8}");
    }

    private readonly bool TryWriteShapeEnd()
    {
        return buffer.TryWrite(
            $"{ShapeAfterAnchor}" +
            $"{reference.Row - 1}" +
            $"{ShapeAfterRow}" +
            $"{reference.Column - 1}" +
            $"{ShapeEnd}");
    }

    private enum Element
    {
        ShapeStart,
        MarginTop,
        ShapeAfterMarginTop,
        Anchor,
        ShapeEnd,
        Done
    }
}
