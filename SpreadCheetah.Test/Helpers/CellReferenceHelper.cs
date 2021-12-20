namespace SpreadCheetah.Test.Helpers;

internal static class CellReferenceHelper
{
    public static string GetExcelColumnName(int columnNumber)
    {
        var columnName = "";

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }
}
