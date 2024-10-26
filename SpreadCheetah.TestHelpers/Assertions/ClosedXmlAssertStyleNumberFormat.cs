using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Helpers;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleNumberFormat(IXLNumberFormat numberFormat)
    : ISpreadsheetAssertStyleNumberFormat
{
    public string? Format => numberFormat.Format is { Length: > 0 } format ? format : null;

    public StandardNumberFormat? StandardFormat =>
        (StandardNumberFormat)numberFormat.NumberFormatId is var format && EnumHelper.IsDefined(format)
            ? format
            : null;
}
