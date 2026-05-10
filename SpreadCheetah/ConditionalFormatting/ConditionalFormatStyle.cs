namespace SpreadCheetah.ConditionalFormatting;

public sealed record ConditionalFormatStyle
{
    public ConditionalFormatBorder Border { get; set; } = new();

    public ConditionalFormatFill Fill { get; set; } = new();

    public ConditionalFormatFont Font { get; set; } = new();

    public string? Format { get; set; } // TODO: Verify in setter.
}
