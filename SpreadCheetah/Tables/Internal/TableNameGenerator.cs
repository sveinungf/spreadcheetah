using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal static class TableNameGenerator
{
    public static string GenerateUniqueTableName(
        HashSet<string> existingTableNames)
    {
        var counter = 1;

        while (true)
        {
            var name = StringHelper.Invariant($"Table{counter}");
            if (!existingTableNames.Contains(name))
                return name;
            counter++;
        }
    }
}
