namespace SpreadCheetah.Styling;

internal sealed record Alignment
{
    public HorizontalAlignment Horizontal { get; set; } // TODO: Validate setter
    public VerticalAlignment Vertical { get; set; } // TODO: Validate setter
    public bool WrapText { get; set; }
    public int Indent { get; set; } // TODO: Validate setter
}
