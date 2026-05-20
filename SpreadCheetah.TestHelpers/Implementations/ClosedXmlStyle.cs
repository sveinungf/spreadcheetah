using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlStyle(IXLStyle style) : IStyle
{
    public IStyleFill Fill => new ClosedXmlStyleFill(style.Fill);

    public IStyleFont Font => new ClosedXmlStyleFont(style.Font);

    public IStyleNumberFormat NumberFormat => new ClosedXmlStyleNumberFormat(style.NumberFormat);
}
