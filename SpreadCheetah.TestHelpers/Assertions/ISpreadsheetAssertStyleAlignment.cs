using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleAlignment
{
    int Indent { get; }
    HorizontalAlignment HorizontalAlignment { get; }
    VerticalAlignment VerticalAlignment { get; }
    bool WrapText { get; }
}
