namespace SpreadCheetah.Images;

internal class ImageCount
{
    public int Jpg { get; private set; }
    public int Png { get; private set; }

    public int TotalCount => Jpg + Png;

    public void Add(ImageType type)
    {
        if (type == ImageType.Jpg)
            ++Jpg;
        else if (type == ImageType.Png)
            ++Png;
    }
}
