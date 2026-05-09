namespace SpreadCheetah.Styling;

public sealed record DifferentialStyle
{
    // TODO: Add border.
    // TODO: Add font.

    public DifferentialFill Fill { get; set; } = new();

    public string? Format { get; set; } // TODO: Verify in setter.
}
