using System;
using System.Collections.Generic;

namespace SpreadCheetah.Worksheets
{
    /// <summary>
    /// Provides options to be used when starting a worksheet with <see cref="Spreadsheet"/>.
    /// </summary>
    public class WorksheetOptions
    {
        internal SortedDictionary<int, ColumnOptions> ColumnOptions { get; } = new SortedDictionary<int, ColumnOptions>();

        /// <summary>
        /// Get options for a column in the worksheet. The first column has column number 1.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public ColumnOptions Column(int columnNumber)
        {
            if (columnNumber < 1) throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column number can't be less than 1.");

            if (!ColumnOptions.TryGetValue(columnNumber, out var options))
            {
                options = new ColumnOptions();
                ColumnOptions.Add(columnNumber, options);
            }

            return options;
        }
    }
}
