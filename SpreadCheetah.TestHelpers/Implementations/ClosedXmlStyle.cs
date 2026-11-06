using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyle : IStyle
{
    public required IStyleAlignment Alignment { get; init; }
    public required IStyleBorder Border { get; init; }
    public required IStyleFill Fill { get; init; }
    public required IStyleFont Font { get; init; }
    public required IStyleNumberFormat NumberFormat { get; init; }

    public static ClosedXmlStyle Create(IXLStyle style)
    {
        return new()
        {
            Alignment = ClosedXmlStyleAlignment.Create(style.Alignment),
            Border = ClosedXmlStyleBorder.Create(style.Border),
            Fill = ClosedXmlStyleFill.Create(style.Fill),
            Font = ClosedXmlStyleFont.Create(style.Font),
            NumberFormat = ClosedXmlStyleNumberFormat.Create(style.NumberFormat)
        };
    }

    public static ClosedXmlStyle Create(Style style)
    {
        return new()
        {
            Alignment = ClosedXmlStyleAlignment.Create(style.Alignment),
            Border = ClosedXmlStyleBorder.Create(style.Border),
            Fill = ClosedXmlStyleFill.Create(style.Fill),
            Font = ClosedXmlStyleFont.Create(style.Font),
            NumberFormat = ClosedXmlStyleNumberFormat.Create(style.Format)
        };
    }
}
