using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleNumberFormat
{
    string? CustomFormat { get; }

    StandardNumberFormat? StandardFormat { get; }
}
