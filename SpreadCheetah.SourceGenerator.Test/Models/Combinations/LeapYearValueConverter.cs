using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Combinations;

public class LeapYearValueConverter : CellValueConverter<int>
{
    public override DataCell ConvertToDataCell(int value)
    {
        return DateTime.IsLeapYear(value)
            ? new DataCell($"{value} (leap year)")
            : new DataCell(value);
    }
}