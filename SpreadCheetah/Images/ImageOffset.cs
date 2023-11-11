namespace SpreadCheetah.Images;

public sealed class ImageOffset
{
    internal int Left { get; private set; }
    internal int Top { get; private set; }
    internal int Right { get; private set; }
    internal int Bottom { get; private set; }

    public static ImageOffset Pixels(int left, int top, int right, int bottom)
    {
        // TODO: Validate arguments. Positive or negative limits?
        return new ImageOffset
        {
            Left = left,
            Top = top,
            Right = right,
            Bottom = bottom
        };
    }
}
