using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class CellValueTypeExtensions
{
    public static string? GetExpectedDefaultNumberFormat(this CellValueType cellValueType) => cellValueType switch
    {
        CellValueType.DateTime => NumberFormats.DateTimeSortable,
        _ => null
    };
}
