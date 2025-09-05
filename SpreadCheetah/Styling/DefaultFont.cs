using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

public sealed record DefaultFont
{
    internal const string DefaultName = "Calibri";
    internal const double DefaultSize = 11;

    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = Guard.FontNameLengthInRange(value);
    }

    private string? _name;

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size
    {
        get => _size;
        set => _size = Guard.FontSizeInRange(value);
    }

    private double _size = DefaultSize;
}
