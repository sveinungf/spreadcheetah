namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal interface ISpreadsheetAssertCell
{
    int? IntValue { get; }
    decimal? DecimalValue { get; }
    string? StringValue { get; }
}
