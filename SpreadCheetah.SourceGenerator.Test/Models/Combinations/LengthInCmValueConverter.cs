using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Combinations;

internal class LengthInCmValueConverter : CellValueConverter<decimal>
{
    public override DataCell ConvertToDataCell(decimal value) => new($"{value} cm");
}
