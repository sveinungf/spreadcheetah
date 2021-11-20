namespace SpreadCheetah.Worksheets;

public class RowOptions
{
    public double? Height
    {
        get => _height;
        set => _height = value is <= 0 or > 409
            ? throw new ArgumentOutOfRangeException(nameof(value), value, "Row height must be between 0 and 409.")
            : value;
    }

    private double? _height;
}
