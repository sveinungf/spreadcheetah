using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using System.Diagnostics;

namespace SpreadCheetah.Helpers;

internal sealed class FileCounter
{
    public int WorksheetsWithImages { get; set; }
    public int WorksheetsWithNotes { get; set; }
    public int TotalTables { get; set; } // TODO: Increment when adding a table
    public int TotalEmbeddedImages { get; private set; }
    public int TotalAddedImages { get; set; }
    public EmbeddedImageTypes EmbeddedImageTypes { get; private set; }

    public void AddEmbeddedImage(ImageType type)
    {
        Debug.Assert(type > ImageType.None && EnumHelper.IsDefined(type));

        if (type == ImageType.Png)
            EmbeddedImageTypes |= EmbeddedImageTypes.Png;
        else if (type == ImageType.Jpeg)
            EmbeddedImageTypes |= EmbeddedImageTypes.Jpeg;

        TotalEmbeddedImages++;
    }
}
