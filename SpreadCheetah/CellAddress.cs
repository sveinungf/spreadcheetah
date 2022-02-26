using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly struct CellAddress : IEquatable<CellAddress>
{
    private static Regex Regex { get; } = new Regex(@"^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$", RegexOptions.None, TimeSpan.FromSeconds(1));

    public string Address { get; }

    private CellAddress(string address) => Address = address;

    public static bool TryCreate(string? value, out CellAddress? address) // TODO: NotNullWhen
    {
        address = null;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (!Regex.IsMatch(value))
            return false;

        address = new CellAddress(value!);
        return true;
    }

    public bool Equals(CellAddress other) => string.Equals(Address, other.Address, StringComparison.Ordinal);
    public override bool Equals(object? obj) => obj is CellAddress other && Equals(other);
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Address);
    public static bool operator ==(in CellAddress left, in CellAddress right) => left.Equals(right);
    public static bool operator !=(in CellAddress left, in CellAddress right) => !left.Equals(right);
}
