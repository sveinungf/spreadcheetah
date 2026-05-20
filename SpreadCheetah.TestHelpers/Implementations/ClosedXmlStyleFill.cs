using ClosedXML.Excel;
using SpreadCheetah.Styling;
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

    public static ClosedXmlStyleFill Create(Fill fill)
    {
        return new()
        {
            Color = fill.Color ?? Color.Transparent
        };
    }
}
