namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetCell
{
    int? IntValue { get; }
    decimal? DecimalValue { get; }
    double? DoubleValue { get; }
    string? StringValue { get; }
    DateTime? DateTimeValue { get; }
    IStyle Style { get; }
    string? Formula { get; }
}
