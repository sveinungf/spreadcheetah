namespace SpreadCheetah.Helpers
{
    internal static class SpreadsheetConstants
    {
        // We are not expecting more than 9,999 sheets in one file
        public const int SheetCountMaxDigits = 4;

        // The maximum number of rows is limited to 1,048,576 in Excel
        public const int RowIndexMaxDigits = 7;

        // The maximum number of unique cell styles is limited to 65,490 in Excel
        public const int StyleIdMaxDigits = 5;
    }
}
