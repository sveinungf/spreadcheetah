using System.Text;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class StringBuilderExtensions
{
    public static StringBuilder AppendLine(this StringBuilder sb, int indentationLevel, string value)
    {
        _ = indentationLevel switch
        {
            <= 0 => sb,
            1 => sb.Append("    "),
            2 => sb.Append("        "),
            3 => sb.Append("            "),
            4 => sb.Append("                "),
            5 => sb.Append("                    "),
            6 => sb.Append("                        "),
            _ => sb.Append(new string(' ', 4 * indentationLevel))
        };

        return sb.AppendLine(value);
    }
}
