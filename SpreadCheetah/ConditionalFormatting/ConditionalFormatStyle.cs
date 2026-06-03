namespace SpreadCheetah.ConditionalFormatting;

public sealed record ConditionalFormatStyle
{
    // TODO: Lazily initialize.
    public ConditionalFormatBorder Border { get; set; } = new();

    public ConditionalFormatFill Fill { get; set; } = new();

    public ConditionalFormatFont Font { get; set; } = new();

    public string? Format { get; set; } // TODO: Verify in setter.

    internal bool IsDefault => this is
    {
        Border.IsDefault: true,
        Fill.IsDefault: true,
        Font.IsDefault: true,
        Format: null
    };
}
