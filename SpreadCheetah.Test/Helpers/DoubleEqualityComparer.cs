namespace SpreadCheetah.Test.Helpers;

internal sealed class DoubleEqualityComparer(double tolerance) : IEqualityComparer<double>
{
    public bool Equals(double x, double y) => Math.Abs(x - y) <= tolerance;
    public int GetHashCode(double obj) => obj.GetHashCode();
}
