using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the alignment parts of a <see cref="Style"/>.
/// </summary>
public sealed record Alignment
{
    /// <summary>Allow text to be shown on multiple lines in the cell.</summary>
    public bool WrapText { get; set; }

    /// <summary>Horizontal alignment for the cell.</summary>
    public HorizontalAlignment Horizontal
    {
        get => _horizontal;
        set
        {
            if (!EnumPolyfill.IsDefined(value))
                ThrowHelper.EnumValueInvalid(nameof(value), value);
            else
                _horizontal = value;
        }
    }

    private HorizontalAlignment _horizontal;

    /// <summary>Vertical alignment for the cell.</summary>
    public VerticalAlignment Vertical
    {
        get => _vertical;
        set
        {
            if (!EnumPolyfill.IsDefined(value))
                ThrowHelper.EnumValueInvalid(nameof(value), value);
            else
                _vertical = value;
        }
    }

    private VerticalAlignment _vertical;

    /// <summary>
    /// Indentation for the cell. The value represents the number of character widths in Excel.
    /// The value can not be negative.
    /// </summary>
    public int Indent
    {
        get => _indent;
        set
        {
            if (value < 0)
                ThrowHelper.ValueIsNegative(nameof(value), value);
            else
                _indent = value;
        }
    }

    private int _indent;
}
