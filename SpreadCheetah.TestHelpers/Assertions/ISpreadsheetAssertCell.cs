namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertCell
{
    int? IntValue { get; }
    decimal? DecimalValue { get; }
    string? StringValue { get; }
}
