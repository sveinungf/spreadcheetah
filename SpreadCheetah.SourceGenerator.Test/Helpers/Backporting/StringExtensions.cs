using System.Text;

namespace SpreadCheetah.SourceGenerator.Test.Helpers.Backporting;

internal static class StringExtensions
{
    public static string ReplaceLineEndings(this string value)
    {
        if (string.Equals(Environment.NewLine, "\n", StringComparison.Ordinal))
        {
#pragma warning disable CA1307 // Specify StringComparison for clarity
            return value.Replace("\r\n", Environment.NewLine);
#pragma warning restore CA1307 // Specify StringComparison for clarity
        }

        var parts = value.Split('\n');
        var sb = new StringBuilder();

        foreach (var part in parts)
        {
            if (part.Length > 1 && part[^1] == '\r')
                sb.Append(part.AsSpan(0, part.Length - 1));
            else
                sb.Append(part);
        }

        return sb.ToString();
    }
}
