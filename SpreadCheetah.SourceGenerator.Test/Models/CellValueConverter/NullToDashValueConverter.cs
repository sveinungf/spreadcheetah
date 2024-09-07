using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

internal sealed class NullToDashValueConverter<T> : CellValueConverter<T?>
{
    public override DataCell ConvertToDataCell(T? value)
    {
        return value is null ? new("-") : new(value.ToString());
    }
}
