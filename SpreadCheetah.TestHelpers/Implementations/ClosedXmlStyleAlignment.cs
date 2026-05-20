using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyleAlignment : IStyleAlignment
{
    public required int Indent { get; init; }
    public required bool WrapText { get; init; }
    public required HorizontalAlignment HorizontalAlignment { get; init; }
    public required VerticalAlignment VerticalAlignment { get; init; }

    public static ClosedXmlStyleAlignment Create(IXLAlignment alignment)
    {
        return new()
        {
            Indent = alignment.Indent,
            WrapText = alignment.WrapText,
            HorizontalAlignment = Map(alignment.Horizontal),
            VerticalAlignment = Map(alignment.Vertical)
        };
    }

    private static HorizontalAlignment Map(XLAlignmentHorizontalValues value)
    {
        return value switch
        {
            XLAlignmentHorizontalValues.General => HorizontalAlignment.None,
            XLAlignmentHorizontalValues.Left => HorizontalAlignment.Left,
            XLAlignmentHorizontalValues.Center => HorizontalAlignment.Center,
            XLAlignmentHorizontalValues.Right => HorizontalAlignment.Right,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, null)
        };
    }

    private static VerticalAlignment Map(XLAlignmentVerticalValues value)
    {
        return value switch
        {
            XLAlignmentVerticalValues.Bottom => VerticalAlignment.Bottom,
            XLAlignmentVerticalValues.Center => VerticalAlignment.Center,
            XLAlignmentVerticalValues.Top => VerticalAlignment.Top,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, null)
        };
    }
}