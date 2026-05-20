using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyleFill : IStyleFill
{
    public required Color Color { get; init; }

    public static ClosedXmlStyleFill Create(IXLFill fill)
    {
        return new()
        {
            Color = fill.BackgroundColor.Color
        };
    }
}
