using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlStyleNumberFormat(IXLNumberFormat numberFormat)
    : IStyleNumberFormat
{
    public string? CustomFormat => numberFormat.Format is { Length: > 0 } format ? format : null;

    public StandardNumberFormat? StandardFormat =>
        (StandardNumberFormat)numberFormat.NumberFormatId is var format && Enum.IsDefined(format)
            ? format
            : null;
}
