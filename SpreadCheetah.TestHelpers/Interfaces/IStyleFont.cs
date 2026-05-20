using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyleFont
{
    bool Bold { get; }
    bool Italic { get; }
    bool Strikethrough { get; }
    double Size { get; }
    string Name { get; }
    Color Color { get; }
    Underline Underline { get; }
}
