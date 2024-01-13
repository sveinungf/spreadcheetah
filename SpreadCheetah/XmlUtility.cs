using SpreadCheetah.Helpers;
using System.Runtime.CompilerServices;
using System.Text;

namespace SpreadCheetah;

public static class XmlUtility
{
    private static readonly UTF8Encoding Utf8NoBom = new(false);

    private static ReadOnlySpan<char> CharsToEscapeSpan =>
    [
        '\x00',
        '\x01',
        '\x02',
        '\x03',
        '\x04',
        '\x05',
        '\x06',
        '\x07',
        '\x08',
        '\x0b',
        '\x0c',
        '\x0e',
        '\x0f',
        '\x10',
        '\x11',
        '\x12',
        '\x13',
        '\x14',
        '\x15',
        '\x16',
        '\x17',
        '\x18',
        '\x19',
        '\x1a',
        '\x1b',
        '\x1c',
        '\x1d',
        '\x1e',
        '\x1f',
        '"',
        '&',
        '\'',
        '<',
        '>'
    ];

#if NET8_0_OR_GREATER
    private static readonly System.Buffers.SearchValues<char> CharsToEscape = System.Buffers.SearchValues.Create(CharsToEscapeSpan);
#else
    private static ReadOnlySpan<char> CharsToEscape => CharsToEscapeSpan;
#endif

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryXmlEncodeToUtf8(ReadOnlySpan<char> source, Span<byte> destination, out int bytesWritten)
    {
        if (!Utf8Helper.DestinationCanFitTranscodedString(source, destination))
        {
            bytesWritten = 0;
            return false;
        }

        var index = source.IndexOfAny(CharsToEscape);
        if (index == -1)
        {
            bytesWritten = Utf8NoBom.GetBytes(source, destination);
            return true;
        }

        bytesWritten = XmlEncodeToUtf8(source, destination, index);
        return true;
    }

    private static int XmlEncodeToUtf8(ReadOnlySpan<char> source, Span<byte> destination, int index)
    {
        var initialDestinationLength = destination.Length;

        while (index != -1)
        {
            if (index > 0)
            {
                var written = Utf8NoBom.GetBytes(source.Slice(0, index), destination);
                source = source.Slice(index);
                destination = destination.Slice(written);
            }

            var replacement = source[0] switch
            {
                '"' => "&quot;"u8,
                '&' => "&amp;"u8,
                '\'' => "&apos;"u8,
                '<' => "&lt;"u8,
                '>' => "&gt;"u8,
                _ => []
            };

            if (replacement.Length > 0)
            {
                replacement.CopyTo(destination);
                destination = destination.Slice(replacement.Length);
            }

            source = source.Slice(1);
            index = source.IndexOfAny(CharsToEscape);
        }

        var finalWritten = Utf8NoBom.GetBytes(source, destination);
        return initialDestinationLength - destination.Length + finalWritten;
    }
}
