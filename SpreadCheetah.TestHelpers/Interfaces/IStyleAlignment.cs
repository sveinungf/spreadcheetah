using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyleAlignment
{
    int Indent { get; }
    HorizontalAlignment HorizontalAlignment { get; }
    VerticalAlignment VerticalAlignment { get; }
    bool WrapText { get; }
}
