namespace SpreadCheetah.ConditionalFormatting;

public sealed record ConditionalFormatStyle
{
    // TODO: Validate that these are not set to null in setters.
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
