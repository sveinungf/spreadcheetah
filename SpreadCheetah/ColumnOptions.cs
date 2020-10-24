using System;

namespace SpreadCheetah.Worksheets
{
    public class ColumnOptions
    {
        /// <summary>
        /// The width of the column. The number represents how many characters can be displayed in the standard font.
        /// Must be between 0 and 255.
        /// </summary>
        public double? Width
        {
            get => _width;
            set => _width = value <= 0 || value > 255
                ? throw new ArgumentOutOfRangeException(nameof(value), value, "Column width must be between 0 and 255.")
                : value;
        }

        private double? _width;
    }
}
