namespace SpreadCheetah.Test.Helpers;

internal static class ColumnName
{
    public static int Parse(string columnName)
    {
        return !SpreadsheetUtility.TryParseColumnName(columnName, out var columnNumber)
            ? throw new ArgumentException("Invalid column name", nameof(columnName))
            : columnNumber;
    }
}
