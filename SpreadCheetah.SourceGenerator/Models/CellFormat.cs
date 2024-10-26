using SpreadCheetah.SourceGenerator.Models.Values;

namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct CellFormat
{
    public string? RawString { get; }
    public StandardNumberFormat? StandardFormat { get; }

    public CellFormat(string rawString) => RawString = rawString;
    public CellFormat(StandardNumberFormat standardFormat) => StandardFormat = standardFormat;
}
