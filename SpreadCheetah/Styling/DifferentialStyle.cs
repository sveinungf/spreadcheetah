namespace SpreadCheetah.Styling;

public sealed record DifferentialStyle
{
    // TODO: Add border.

    public DifferentialFill Fill { get; set; } = new();

    public DifferentialFont Font { get; set; } = new();

    public string? Format { get; set; } // TODO: Verify in setter.
}
