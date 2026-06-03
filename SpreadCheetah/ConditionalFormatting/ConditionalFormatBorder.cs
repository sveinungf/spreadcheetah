namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents all border parts of a <see cref="ConditionalFormatStyle"/>.
/// </summary>
public sealed record ConditionalFormatBorder
{
    /// <summary>Left border for the cell.</summary>
    public ConditionalFormatEdgeBorder Left
    {
        get => _left ??= new();
        set => _left = value;
    }

    private ConditionalFormatEdgeBorder? _left;
    internal ConditionalFormatEdgeBorder? GetLeftOrDefault() => _left;

    /// <summary>Right border for the cell.</summary>
    public ConditionalFormatEdgeBorder Right
    {
        get => _right ??= new();
        set => _right = value;
    }

    private ConditionalFormatEdgeBorder? _right;
    internal ConditionalFormatEdgeBorder? GetRightOrDefault() => _right;

    /// <summary>Top border for the cell.</summary>
    public ConditionalFormatEdgeBorder Top
    {
        get => _top ??= new();
        set => _top = value;
    }

    private ConditionalFormatEdgeBorder? _top;
    internal ConditionalFormatEdgeBorder? GetTopOrDefault() => _top;

    /// <summary>Bottom border for the cell.</summary>
    public ConditionalFormatEdgeBorder Bottom
    {
        get => _bottom ??= new();
        set => _bottom = value;
    }

    private ConditionalFormatEdgeBorder? _bottom;
    internal ConditionalFormatEdgeBorder? GetBottomOrDefault() => _bottom;

    internal bool IsDefault => this is
    {
        _left: null or { IsDefault: true },
        _right: null or { IsDefault: true },
        _top: null or { IsDefault: true },
        _bottom: null or { IsDefault: true }
    };
}
