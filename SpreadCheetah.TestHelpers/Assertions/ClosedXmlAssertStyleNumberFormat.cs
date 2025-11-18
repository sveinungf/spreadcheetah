using ClosedXML.Excel;
using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleNumberFormat(IXLNumberFormat numberFormat)
    : ISpreadsheetAssertStyleNumberFormat
{
    public string? CustomFormat => numberFormat.Format is { Length: > 0 } format ? format : null;

    public StandardNumberFormat? StandardFormat =>
        (StandardNumberFormat)numberFormat.NumberFormatId is var format && Enum.IsDefined(format)
            ? format
            : null;
}
