namespace SpreadCheetah.Tables.Internal;

internal static class TableDictionaryExtensions
{
    public static WorksheetTableInfo? GetActive(
        this Dictionary<string, WorksheetTableInfo>? tables) => tables.GetActive(out _);

    public static WorksheetTableInfo? GetActive(
        this Dictionary<string, WorksheetTableInfo>? tables,
        out bool multipleActiveTables)
    {
        multipleActiveTables = false;

        if (tables is null)
            return null;

        WorksheetTableInfo? result = null;
        var activeTables = 0;

        foreach (var table in tables.Values)
        {
            if (!table.Active)
                continue;

            result = table;
            activeTables++;

            if (activeTables > 1)
            {
                multipleActiveTables = true;
                return null;
            }
        }

        return result;
    }
}
