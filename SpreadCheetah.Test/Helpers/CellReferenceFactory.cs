namespace SpreadCheetah.Test.Helpers;

internal static class CellReferenceFactory
{
    public static List<string> RowReferences(int rowNumber, int count)
    {
        return Enumerable.Range(1, count)
            .Select(x => SpreadsheetUtility.GetColumnName(x) + rowNumber)
            .ToList();
    }
}
