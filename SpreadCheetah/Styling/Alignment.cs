using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

internal sealed record Alignment
{
    public bool WrapText { get; set; }

    public HorizontalAlignment Horizontal
    {
        get => _horizontal;
        set
        {
            if (!EnumHelper.IsDefined(value))
                ThrowHelper.EnumValueInvalid(nameof(value), value);
            else
                _horizontal = value;
        }
    }

    private HorizontalAlignment _horizontal;

    public VerticalAlignment Vertical
    {
        get => _vertical;
        set
        {
            if (!EnumHelper.IsDefined(value))
                ThrowHelper.EnumValueInvalid(nameof(value), value);
            else
                _vertical = value;
        }
    }

    private VerticalAlignment _vertical;

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
