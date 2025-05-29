namespace SpreadCheetah.Worksheets.Internal;

internal sealed class WorksheetRowHeightRun(uint startingRowIndex, double height)
{
    public uint StartingRowIndex { get; } = startingRowIndex;
    public double Height { get; } = height;
    public uint RowCount { get; private set; } = 1;

    public bool TryContinueWith(uint rowIndex, double height)
    {
        var expectedRowIndex = StartingRowIndex + RowCount;
        if (rowIndex != expectedRowIndex)
            return false;

        // Using a tolerance for floating-point comparison
        if (Math.Abs(height - Height) > 0.0001)
            return false;

        RowCount++;
        return true;
    }
}
