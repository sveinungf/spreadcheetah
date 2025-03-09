namespace SpreadCheetah.Tables.Internal;

internal static class TableDictionaryExtensions
{
    public static WorksheetTableInfo? GetActive(
        this List<WorksheetTableInfo>? tables)
    {
        return tables?.Find(x => x.Active);
    }
}
