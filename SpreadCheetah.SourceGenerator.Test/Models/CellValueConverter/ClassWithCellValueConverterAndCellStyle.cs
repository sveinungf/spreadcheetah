using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

internal class ClassWithCellValueConverterAndCellStyle
{
    [CellStyle("ID style")]
    [CellValueConverter(typeof(UpperCaseValueConverter))]
    public required string Id { get; set; }
}
