using SpreadCheetah.Images;

namespace SpreadCheetah.Helpers;

internal sealed class FileCounter
{
    public int WorksheetsWithImages { get; set; }
    public int WorksheetsWithNotes { get; set; }
    public int Jpg { get; private set; } // TODO: Maybe this could be a flag enum?
    public int Png { get; private set; } // TODO: Maybe this could be a flag enum?

    public int TotalImageCount => Jpg + Png;

    public void AddImage(ImageType type)
    {
        if (type == ImageType.Jpg)
            ++Jpg;
        else if (type == ImageType.Png)
            ++Png;
    }
}
