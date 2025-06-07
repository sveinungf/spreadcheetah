using System.Globalization;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class IntegerExtensions
{
    /// <summary>EMU = English Metric Unit</summary>
    public static int PixelsToEmu(this int n) => n * 9525; // TODO: Check if this is correct for height

    public static int PixelsToWidthOffset(this int n) => n * 7620; // TODO: Same for height?

    public static string ToStringInvariant(this int n) => n.ToString(CultureInfo.InvariantCulture);

    public static void EnsureValidImageDimension(this int value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value <= 0)
            ThrowHelper.ImageDimensionZeroOrNegative(paramName, value);
        if (value > SpreadsheetConstants.MaxImageDimension)
            ThrowHelper.ImageDimensionTooLarge(paramName, value);
    }

    public static (int Width, int Height) Scale(this (int Width, int Height) dimensions, float scale)
    {
        var width = (int)Math.Round(dimensions.Width * scale, MidpointRounding.AwayFromZero);
        var height = (int)Math.Round(dimensions.Height * scale, MidpointRounding.AwayFromZero);
        return (width, height);
    }
}
