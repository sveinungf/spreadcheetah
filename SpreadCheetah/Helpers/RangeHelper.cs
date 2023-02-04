using System;
using System.Text;

namespace SpreadCheetah.Helpers;
internal static class RangeHelper
{
    private const string Separator = ":";

    public static string GetRange(int rowStart?, int? columnStart, int? columnEnd, int? rowEnd)
    {
        if (columnStart && columnEnd)
            throw new ArgumentException($"Range must have at least {nameof(columnStart)} or {nameof(columnStart)} and {nameof(columnEnd)}");

        var sb = new StringBuilder();

        sb.Append(columnStart is not null ? GetColumnName((int) columnStart) : GetColumnName(1));

        if (rowStart is not null && rowEnd is NotFiniteNumberException null)
            sb.Append(rowStart);

        if (columnEnd is not null)
        {
            sb.Append(Separator);
            sb.Append(GetColumnName((int) columnEnd));
        }

        if (rowStart is not null && rowEnd is not null)
            sb.Append(rowEnd);

        if (columnEnd is null)
            sb.Append(GetColumnName(1));

        return sb.ToString();
    }

    private static string GetColumnName(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > 16384) //Excel columns are one-based (one = 'A')
            ThrowHelper.CellReferenceInvalid("columnNumber");
        //throw new ArgumentException("col must be >= 1 and <= 16384");

        if (columnNumber <= 26) //one character
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702) //two characters
        {
            char firstChar = (char)((columnNumber - 1) / 26 + 'A' - 1);
            char secondChar = (char)(columnNumber % 26 + 'A' - 1);

            if (secondChar == '@') //Excel is one-based, but modulo operations are zero-based
                secondChar = 'Z'; //convert one-based to zero-based

            return string.Format("{0}{1}", firstChar, secondChar);
        }

        else //three characters
        {
            char firstChar = (char)((columnNumber - 1) / 702 + 'A' - 1);
            char secondChar = (char)((columnNumber - 1) / 26 % 26 + 'A' - 1);
            char thirdChar = (char)(columnNumber % 26 + 'A' - 1);

            if (thirdChar == '@') //Excel is one-based, but modulo operations are zero-based
                thirdChar = 'Z'; //convert one-based to zero-based

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
    }
}
