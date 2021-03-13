using System.Globalization;
using System.Text;

namespace SpreadCheetah.Helpers
{
    internal static class StringBuilderExtensions
    {
        public static StringBuilder AppendDouble(this StringBuilder sb, double value)
        {
            return sb.AppendFormat(CultureInfo.InvariantCulture, "{0:G15}", value);
        }
    }
}
