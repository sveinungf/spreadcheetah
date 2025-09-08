using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleFont
{
    bool Bold { get; }
    bool Italic { get; }
    bool Strikethrough { get; }
    string Name { get; }
    Color Color { get; }
    Underline Underline { get; }
}
