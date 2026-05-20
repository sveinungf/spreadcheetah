namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyle
{
    IStyleFill Fill { get; }
    IStyleFont Font { get; }
    IStyleNumberFormat NumberFormat { get; }
}
