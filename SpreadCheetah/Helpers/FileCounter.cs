using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using System.Diagnostics;

namespace SpreadCheetah.Helpers;

internal sealed class FileCounter
{
    public int WorksheetsWithImages { get; set; }
    public int WorksheetsWithNotes { get; set; }
    public int TotalImageCount { get; private set; }
    public AddedImageTypes AddedImageTypes { get; private set; }

    public void AddImage(ImageType type)
    {
        Debug.Assert(type > ImageType.None && EnumHelper.IsDefined(type));

        if (type == ImageType.Png)
            AddedImageTypes |= AddedImageTypes.Png;
        else if (type == ImageType.Jpeg)
            AddedImageTypes |= AddedImageTypes.Jpeg;

        TotalImageCount++;
    }
}
