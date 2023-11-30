using SpreadCheetah.Helpers;

namespace SpreadCheetah.Images.Internal;

internal static class ImageValidator
{
    public static void EnsureValidCanvas(ImageCanvas canvas, EmbeddedImage image)
    {
        if (canvas.Options.HasFlag(ImageCanvasOptions.Scaled))
        {
            var scale = canvas.ScaleValue;
            var originalDimensions = (image.Width, image.Height);
            var (width, height) = originalDimensions.Scale(scale);
            width.EnsureValidImageDimension();
            height.EnsureValidImageDimension();
        }
    }
}
