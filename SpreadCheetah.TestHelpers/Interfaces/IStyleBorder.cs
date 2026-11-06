using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyleBorder
{
    Color BottomColor { get; }
    BorderStyle BottomStyle { get; }
    Color DiagonalColor { get; }
    BorderStyle DiagonalStyle { get; }
    Color LeftColor { get; }
    BorderStyle LeftStyle { get; }
    Color RightColor { get; }
    BorderStyle RightStyle { get; }
    Color TopColor { get; }
    BorderStyle TopStyle { get; }
}
