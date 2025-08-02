using ClosedXML.Excel;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleFill(IXLFill fill) : ISpreadsheetAssertStyleFill
{
    public Color Color => fill.BackgroundColor.Color;
}
