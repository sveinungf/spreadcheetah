namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyle
{
    IStyleAlignment Alignment { get; }
    IStyleBorder Border { get; }
    IStyleFill Fill { get; }
    IStyleFont Font { get; }
    IStyleNumberFormat NumberFormat { get; }
}
