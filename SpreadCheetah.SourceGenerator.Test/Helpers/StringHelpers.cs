namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal static class StringHelpers
{
    public static string ReplaceLineEndings(string value)
    {
#if NET6_0_OR_GREATER
        return value.ReplaceLineEndings();
#else
        if (string.Equals(Environment.NewLine, "\n", StringComparison.Ordinal))
#pragma warning disable CA1307 // Specify StringComparison for clarity
            return value.Replace("\r\n", Environment.NewLine);

        var parts = value.Split('\n');
        var sb = new System.Text.StringBuilder();

        foreach (var part in parts)
        {
            if (part.Length > 1 && part[^1] == '\r')
                sb.Append(part.AsSpan(0, part.Length - 1));
            else
                sb.Append(part);
        }

        return sb.ToString();
#endif
    }
}
