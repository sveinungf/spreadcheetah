using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly record struct CellReference
{
    private static Regex Regex { get; } = new Regex(@"^\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$", RegexOptions.None, TimeSpan.FromSeconds(1));

    public string Reference { get; }

    private CellReference(string reference) => Reference = reference;

    private static bool TryCreate(string value, [NotNullWhen(true)] out CellReference? reference)
    {
        reference = null;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (!Regex.IsMatch(value))
            return false;

        reference = new CellReference(value!);
        return true;
    }

    public static CellReference Create(string value, [CallerArgumentExpression("value")] string? paramName = null)
    {
        if (!TryCreate(value, out var reference))
            ThrowHelper.CellReferenceInvalid(paramName);

        return reference.Value;
    }
}
