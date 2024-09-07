using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

public class UpperCaseValueConverter : CellValueConverter<string?>
{
    public override DataCell ConvertToDataCell(string? value) => new(value?.ToUpperInvariant());
}
