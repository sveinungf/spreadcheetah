namespace SpreadCheetah.Styling.Internal;

internal sealed class DefaultStyling(int? initialDateTimeStyleId)
{
    private readonly int? _initialDateTimeStyleId = initialDateTimeStyleId;
    private uint? _currentRowIndex;

    public int? DateTimeStyleId { get; private set; } = initialDateTimeStyleId;

    public void SetRowDateTimeStyleId(uint rowIndex, int? styleId)
    {
        _currentRowIndex = rowIndex;
        DateTimeStyleId = styleId;
    }

    public void ClearRowDateTimeStyleId()
    {
        if (_currentRowIndex.HasValue)
        {
            _currentRowIndex = null;
            DateTimeStyleId = _initialDateTimeStyleId;
        }
    }
}