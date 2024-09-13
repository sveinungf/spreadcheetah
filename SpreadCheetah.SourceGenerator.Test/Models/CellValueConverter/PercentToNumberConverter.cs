using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

public class PercentToNumberConverter : CellValueConverter<Percent?>
{
    public override DataCell ConvertToDataCell(Percent? value)
    {
        return value is null ? new("-") : new(value.GetValueOrDefault().Value);
    }
}