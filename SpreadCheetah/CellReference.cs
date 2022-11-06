using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly record struct CellReference
{
    private static Regex Regex { get; } = new Regex(@"^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$", RegexOptions.None, TimeSpan.FromSeconds(1));

    public string Reference { get; }

    private CellReference(string reference) => Reference = reference;

    public static bool TryCreate(string? value, [NotNullWhen(true)] out CellReference? reference)
    {
        reference = null;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (!Regex.IsMatch(value))
            return false;

        reference = new CellReference(value!);
        return true;
    }
}
