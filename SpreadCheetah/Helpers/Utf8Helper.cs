using System.Runtime.CompilerServices;
using System.Text;

namespace SpreadCheetah.Helpers;

internal static class Utf8Helper
{
    public const int MaxBytePerChar = 6;

    public static readonly UTF8Encoding Utf8NoBom = new(false);

    public static int GetBytes(ReadOnlySpan<char> chars, Span<byte> bytes) => Utf8NoBom.GetBytes(chars, bytes);

    public static bool TryGetBytes(ReadOnlySpan<char> chars, Span<byte> bytes, out int bytesWritten)
    {
        // Try first with an approximate value length, then try with a more accurate value length
        if (DestinationCanFitTranscodedString(chars, bytes))
        {
            bytesWritten = Utf8NoBom.GetBytes(chars, bytes);
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool DestinationCanFitTranscodedString(ReadOnlySpan<char> chars, Span<byte> bytes)
    {
        return bytes.Length >= chars.Length * MaxBytePerChar || bytes.Length >= Utf8NoBom.GetByteCount(chars);
    }
}
