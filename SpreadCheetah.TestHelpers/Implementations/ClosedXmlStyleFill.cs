using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlStyleFill(IXLFill fill) : IStyleFill
{
    public Color Color => fill.BackgroundColor.Color;
}
