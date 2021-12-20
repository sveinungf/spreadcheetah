using System.Globalization;
using System.Text;

namespace SpreadCheetah.Helpers;

internal static class StringBuilderExtensions
{
    public static StringBuilder AppendDouble(this StringBuilder sb, double value)
    {
        return sb.AppendFormat(CultureInfo.InvariantCulture, "{0:G15}", value);
    }

    public static StringBuilder AppendColumnName(this StringBuilder sb, int columnNumber)
    {
        var columnName = "";

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return sb.Append(columnName);
    }
}
