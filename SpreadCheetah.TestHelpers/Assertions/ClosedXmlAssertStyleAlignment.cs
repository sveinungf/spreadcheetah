using ClosedXML.Excel;
using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleAlignment(IXLAlignment alignment)
    : ISpreadsheetAssertStyleAlignment
{
    public int Indent => alignment.Indent;

    public HorizontalAlignment HorizontalAlignment => alignment.Horizontal switch
    {
        XLAlignmentHorizontalValues.General => HorizontalAlignment.None,
        XLAlignmentHorizontalValues.Left => HorizontalAlignment.Left,
        XLAlignmentHorizontalValues.Center => HorizontalAlignment.Center,
        XLAlignmentHorizontalValues.Right => HorizontalAlignment.Right,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment.Horizontal, null)
    };

    public VerticalAlignment VerticalAlignment => alignment.Vertical switch
    {
        XLAlignmentVerticalValues.Bottom => VerticalAlignment.Bottom,
        XLAlignmentVerticalValues.Center => VerticalAlignment.Center,
        XLAlignmentVerticalValues.Top => VerticalAlignment.Top,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment.Vertical, null)
    };

    public bool WrapText => alignment.WrapText;
}
