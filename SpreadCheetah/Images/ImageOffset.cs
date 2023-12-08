using System.Runtime.InteropServices;

namespace SpreadCheetah.Images;

/// <summary>
/// Represents the offset part of <see cref="ImageOptions"/>.
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly record struct ImageOffset
{
    internal int Left { get; private init; }
    internal int Top { get; private init; }
    internal int Right { get; private init; }
    internal int Bottom { get; private init; }

    /// <summary>
    /// Defines the offset of each side of the image in pixels.
    /// </summary>
    public static ImageOffset Pixels(int left, int top, int right, int bottom)
    {
        return new ImageOffset
        {
            Left = left,
            Top = top,
            Right = right,
            Bottom = bottom
        };
    }
}
