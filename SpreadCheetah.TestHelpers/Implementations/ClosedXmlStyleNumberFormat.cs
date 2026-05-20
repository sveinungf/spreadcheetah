using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyleNumberFormat : IStyleNumberFormat
{
    public required string? CustomFormat { get; init; }
    public required StandardNumberFormat? StandardFormat { get; init; }

    public static ClosedXmlStyleNumberFormat Create(IXLNumberFormat numberFormat)
    {
        return new()
        {
            CustomFormat = numberFormat.Format is { Length: > 0 } format ? format : null,
            StandardFormat = Map(numberFormat.NumberFormatId)
        };
    }

    private static StandardNumberFormat? Map(int formatId)
    {
        var standardFormat = (StandardNumberFormat)formatId;
        return Enum.IsDefined(standardFormat) ? standardFormat : null;
    }
}
