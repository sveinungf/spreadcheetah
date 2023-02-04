using System.Text;

namespace SpreadCheetah.Helpers;
internal static class RangeHelper
{
    public static string GetRange(int? rowStart, int columnStart, int? columnEnd, int? rowEnd)
    {
        var sb = new StringBuilder();

        sb.Append(GetColumnName(columnStart));

        if (rowStart is not null && rowEnd is not null)
            sb.Append(rowStart);

        if (columnEnd is not null)
        {
            sb.Append(':');
            sb.Append(GetColumnName(columnEnd.Value));
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
            ThrowHelper.CellReferenceInvalid(nameof(columnNumber));

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
