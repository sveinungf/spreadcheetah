using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal static class TableNameGenerator
{
    public static string GenerateUniqueName(
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
}
