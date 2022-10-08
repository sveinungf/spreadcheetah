using System.Buffers.Text;
using System.Diagnostics;
using System.Text;

namespace SpreadCheetah.Helpers;

internal static class Utf8Helper
{
    public const int MaxBytePerChar = 6;

    private static readonly UTF8Encoding Utf8NoBom = new(false);

    public static int GetBytes(int number, Span<byte> destination)
    {
        Utf8Formatter.TryFormat(number, destination, out var bytesWritten);
        return bytesWritten;
    }

    public static int GetBytes(double number, Span<byte> destination)
    {
        Utf8Formatter.TryFormat(number, destination, out var bytesWritten);
        return bytesWritten;
    }

    public static int GetBytes(string chars, Span<byte> bytes, bool assertSize = true) => GetBytes(chars.AsSpan(), bytes, assertSize);

    public static int GetBytes(ReadOnlySpan<char> chars, Span<byte> bytes, bool assertSize = true)
    {
        if (assertSize)
            Debug.Assert(Utf8NoBom.GetByteCount(chars) <= SpreadCheetahOptions.MinimumBufferSize);

        return Utf8NoBom.GetBytes(chars, bytes);
    }

    public static byte[] GetBytes(string s) => Utf8NoBom.GetBytes(s);
    public static int GetByteCount(string chars) => Utf8NoBom.GetByteCount(chars);

    public static bool TryGetBytes(ReadOnlySpan<char> chars, Span<byte> bytes, out int bytesWritten)
    {
        return Utf8NoBom.TryGetBytes(chars, bytes, out bytesWritten);
    }
}
