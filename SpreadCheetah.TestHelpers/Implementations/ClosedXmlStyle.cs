using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlStyle(IXLStyle style) : IStyle
{
    public IStyleAlignment Alignment => ClosedXmlStyleAlignment.Create(style.Alignment);
    public IStyleBorder Border => ClosedXmlStyleBorder.Create(style.Border);
    public IStyleFill Fill => new ClosedXmlStyleFill(style.Fill);
    public IStyleFont Font => new ClosedXmlStyleFont(style.Font);
    public IStyleNumberFormat NumberFormat => new ClosedXmlStyleNumberFormat(style.NumberFormat);
}
