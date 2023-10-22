namespace SpreadCheetah.Images;

internal sealed class Counter
{
    public int Jpg { get; private set; }
    public int Png { get; private set; }

    public int TotalImageCount => Jpg + Png;

    public void AddImage(ImageType type)
    {
        if (type == ImageType.Jpg)
            ++Jpg;
        else if (type == ImageType.Png)
            ++Png;
    }
}
