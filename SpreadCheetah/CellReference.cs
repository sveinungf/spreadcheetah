using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly struct CellReference : IEquatable<CellReference>
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

    public bool Equals(CellReference other) => string.Equals(Reference, other.Reference, StringComparison.Ordinal);
    public override bool Equals(object? obj) => obj is CellReference other && Equals(other);
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Reference);
    public static bool operator ==(in CellReference left, in CellReference right) => left.Equals(right);
    public static bool operator !=(in CellReference left, in CellReference right) => !left.Equals(right);
}
