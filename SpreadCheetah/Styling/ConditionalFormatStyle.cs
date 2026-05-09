namespace SpreadCheetah.Styling;

public sealed record ConditionalFormatStyle
{
    // TODO: Add border.

    public ConditionalFormatFill Fill { get; set; } = new();

    public ConditionalFormatFont Font { get; set; } = new();

    public string? Format { get; set; } // TODO: Verify in setter.
}
