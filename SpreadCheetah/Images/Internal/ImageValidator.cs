using SpreadCheetah.Helpers;

namespace SpreadCheetah.Images.Internal;

internal static class ImageValidator
{
    public static void EnsureValidCanvas(ImageCanvas canvas, EmbeddedImage image)
    {
        var originalDimensions = (image.Width, image.Height);
        if (canvas.Options.HasFlag(ImageCanvasOptions.Scale))
        {
            var scale = canvas.ScaleValue;
            var (width, height) = originalDimensions.Scale(scale);
            width.EnsureValidImageDimension(nameof(canvas));
            height.EnsureValidImageDimension(nameof(canvas));
        }
    }
}
