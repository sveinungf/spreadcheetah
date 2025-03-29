using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleFont
{
    bool Bold { get; }
    bool Italic { get; }
    bool Strikethrough { get; }
    Underline Underline { get; }
}
