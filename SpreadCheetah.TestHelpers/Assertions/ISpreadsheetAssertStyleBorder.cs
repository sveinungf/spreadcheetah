using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyleBorder
{
    Color BottomColor { get; }
    BorderStyle BottomStyle { get; }
}
