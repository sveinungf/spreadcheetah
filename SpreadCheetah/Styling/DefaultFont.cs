using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

public sealed class DefaultFont
{
    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = Guard.MaxLength(value, 31);
    }

    private string? _name;

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size
    {
        get => _size;
        set => _size = Guard.FontSizeInRange(value);
    }

    private double _size = 11;
}
