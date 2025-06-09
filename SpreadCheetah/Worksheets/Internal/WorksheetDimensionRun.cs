namespace SpreadCheetah.Worksheets.Internal;

internal sealed class WorksheetDimensionRun(uint startingIndex, double dimension)
{
    public uint StartingIndex { get; } = startingIndex;
    public double Dimension { get; } = dimension;
    public uint Count { get; private set; } = 1;
    public uint LastIndex => StartingIndex + Count - 1;

    public bool ContainsIndex(uint index)
    {
        return index >= StartingIndex && index <= LastIndex;
    }

    public bool TryContinueWith(uint index, double dimension)
    {
        var expectedIndex = StartingIndex + Count;
        if (index != expectedIndex)
            return false;

        // Using a tolerance for floating-point comparison
        if (Math.Abs(dimension - Dimension) > 0.0001)
            return false;

        Count++;
        return true;
    }
}
