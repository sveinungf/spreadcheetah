using System;
using System.Buffers.Text;
using System.Diagnostics;
using System.Text;

namespace SpreadCheetah.Helpers
{
    internal static class Utf8Helper
    {
        public const int MaxBytePerChar = 6;

        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        public static int GetBytes(int number, Span<byte> destination)
        {
            Utf8Formatter.TryFormat(number, destination, out var bytesWritten);
            return bytesWritten;
        }

#if NETSTANDARD2_0
        public static unsafe int GetBytes(string chars, Span<byte> bytes, bool assertSize = true)
        {
            fixed (char* charsPointer = chars)
            fixed (byte* bytesPointer = bytes)
            {
                if (assertSize)
                    Debug.Assert(Utf8NoBom.GetByteCount(charsPointer, chars.Length) <= SpreadCheetahOptions.MinimumBufferSize);

                return Utf8NoBom.GetBytes(charsPointer, chars.Length, bytesPointer, bytes.Length);
            }
        }

        public static unsafe int GetBytes(ReadOnlySpan<char> chars, Span<byte> bytes, bool assertSize = true)
        {
            fixed (char* charsPointer = chars)
            fixed (byte* bytesPointer = bytes)
            {
                if (assertSize)
                    Debug.Assert(Utf8NoBom.GetByteCount(charsPointer, chars.Length) <= SpreadCheetahOptions.MinimumBufferSize);

                return Utf8NoBom.GetBytes(charsPointer, chars.Length, bytesPointer, bytes.Length);
            }
        }
#else
        public static int GetBytes(ReadOnlySpan<char> chars, Span<byte> bytes, bool assertSize = true)
        {
            if (assertSize)
                Debug.Assert(Utf8NoBom.GetByteCount(chars) <= SpreadCheetahOptions.MinimumBufferSize);

            return Utf8NoBom.GetBytes(chars, bytes);
        }
#endif

        public static byte[] GetBytes(string s) => Utf8NoBom.GetBytes(s);
        public static int GetByteCount(string chars) => Utf8NoBom.GetByteCount(chars);
    }
}
