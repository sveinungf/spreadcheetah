namespace SpreadCheetah.Images;

internal static class FileSignature
{
    // Will get an unwanted byte at the start if attempting to use a UTF-8 string literal here, e.g. "\u0089PNG\r\n\u001A\n"u8
    private static ReadOnlySpan<byte> PngHeader => new[] { (byte)0x89, (byte)'P', (byte)'N', (byte)'G', (byte)'\r', (byte)'\n', (byte)0x1A, (byte)'\n' };

    public static ImageType? GetImageTypeFromHeader(ReadOnlySpan<byte> header)
    {
        if (header.StartsWith(PngHeader))
            return ImageType.Png;

        return null;
    }
}
