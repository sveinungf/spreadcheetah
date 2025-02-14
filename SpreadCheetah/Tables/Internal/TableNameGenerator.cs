using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal static class TableNameGenerator
{
    public static string GenerateUniqueTableName(
        Dictionary<string, WorksheetTableInfo> existingTables)
    {
        var counter = 1;

        while (true)
        {
            var name = StringHelper.Invariant($"Table{counter}");
            if (!existingTables.ContainsKey(name))
                return name;
            counter++;
        }
    }

    public static string GenerateUniqueHeaderName(
        HashSet<string> existingColumnNames)
    {
        var counter = 1;

        while (true)
        {
            var name = StringHelper.Invariant($"Column{counter}");
            if (!existingColumnNames.Contains(name))
                return name;
            counter++;
        }
    }
}
