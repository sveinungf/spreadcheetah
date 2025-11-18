#if NETSTANDARD2_0
namespace System.Text;

internal static class Utf8EncodingExtensions
{
    public static unsafe int GetBytes(this UTF8Encoding encoding, ReadOnlySpan<char> chars, Span<byte> bytes)
    {
        if (chars.IsEmpty) return 0;

        fixed (char* charsPointer = chars)
        fixed (byte* bytesPointer = bytes)
        {
            return encoding.GetBytes(charsPointer, chars.Length, bytesPointer, bytes.Length);
        }
    }

    public static unsafe int GetByteCount(this UTF8Encoding encoding, ReadOnlySpan<char> chars)
    {
        if (chars.IsEmpty) return 0;

        fixed (char* charsPointer = chars)
        {
            return encoding.GetByteCount(charsPointer, chars.Length);
        }
    }
}
#endif
