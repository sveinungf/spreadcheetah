using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleNumberFormat
{
    string? Format { get; }

    StandardNumberFormat? StandardFormat { get; }
}
