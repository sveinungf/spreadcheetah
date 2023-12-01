namespace SpreadCheetah.Images;

public readonly record struct ImageOffset
{
    internal int Left { get; private init; }
    internal int Top { get; private init; }
    internal int Right { get; private init; }
    internal int Bottom { get; private init; }

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
