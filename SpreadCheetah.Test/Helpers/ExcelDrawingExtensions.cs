using OfficeOpenXml.Drawing;

namespace SpreadCheetah.Test.Helpers;

internal static class ExcelDrawingExtensions
{
    public static (int Width, int Height) GetActualDimensions(this ExcelDrawing drawing)
    {
        var widthObject = ReflectionHelper.GetInstanceField(drawing, "_width");
        if (widthObject is not int width)
            throw new ArgumentException("Could not get width.", nameof(drawing), null);

        var heightObject = ReflectionHelper.GetInstanceField(drawing, "_height");
        if (heightObject is not int height)
            throw new ArgumentException("Could not get height.", nameof(drawing), null);

        return (width, height);
    }
}
