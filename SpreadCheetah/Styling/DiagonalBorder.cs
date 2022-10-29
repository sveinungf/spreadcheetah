using System.Drawing;

namespace SpreadCheetah.Styling;

internal class DiagonalBorder
{
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;
    public Color? Color { get; set; }
    public DiagonalBorderType Type { get; set; } = DiagonalBorderType.None;
}