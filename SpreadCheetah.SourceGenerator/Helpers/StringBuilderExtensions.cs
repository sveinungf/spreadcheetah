using System.Text;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class StringBuilderExtensions
{
    public static StringBuilder AppendLine(this StringBuilder sb, int indentationLevel, string value)
    {
        return sb.Append(new string(' ', 4 * indentationLevel)).AppendLine(value);
    }
}
