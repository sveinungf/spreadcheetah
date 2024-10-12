using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

internal class ClassWithCellValueConverter
{
    [CellValueConverter<UpperCaseValueConverter>]
    public string? Name { get; set; }
}
