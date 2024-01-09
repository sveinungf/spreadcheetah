using SpreadCheetah.Helpers;
using System.Buffers;
#if NET8_0_OR_GREATER
using System.Collections.Frozen;
#endif
using System.Text;

namespace SpreadCheetah;

public static class XmlUtility
{
#if !NET8_0_OR_GREATER
    public static string XmlEncode(string value) => "";
    public static bool TryXmlEncodeToUtf8(ReadOnlySpan<char> source, Span<byte> destination, out int bytesWritten) { bytesWritten = 0; return false; }
#else
    private static readonly SearchValues<char> SearchCharacters;
    private static readonly FrozenDictionary<char, string> Replacements;
    private static readonly FrozenDictionary<char, byte[]> ByteReplacements;

    static XmlUtility()
    {
        List<KeyValuePair<char, string>> replacements =
        [
            new('<', "&lt;"),
            new('>', "&gt;"),
            new('&', "&amp;"),
            new('"', "&quot;"),
            new('\'', "&apos;"),
        ];

        for (var i = 0; i < 32; ++i)
        {
            var character = (char)i;
            if (character is not ('\t' or '\n' or '\r'))
                replacements.Add(new KeyValuePair<char, string>(character, ""));
        }

        Replacements = replacements.ToFrozenDictionary();
        ByteReplacements = replacements.Select(x => KeyValuePair.Create(x.Key, Utf8NoBom.GetBytes(x.Value))).ToFrozenDictionary();
        SearchCharacters = SearchValues.Create(replacements.Select(x => x.Key).ToArray());
    }

    public static string XmlEncode(string value)
    {
        var index = value.AsSpan().IndexOfAny(SearchCharacters);
        if (index == -1)
            return value;

        var remainingValue = value.AsSpan();
        var sb = new StringBuilder();
        while (index != -1)
        {
            if (index > 0)
                sb.Append(remainingValue.Slice(0, index));

            var characterToReplace = remainingValue[index];
            var replacement = Replacements[characterToReplace];
            sb.Append(replacement);

            remainingValue = remainingValue.Slice(index + 1);
            index = remainingValue.IndexOfAny(SearchCharacters);
        }

        sb.Append(remainingValue);
        return sb.ToString();
    }

    private static readonly UTF8Encoding Utf8NoBom = new(false);

    public static bool TryXmlEncodeToUtf8(ReadOnlySpan<char> source, Span<byte> destination, out int bytesWritten)
    {
        if (!Utf8Helper.DestinationCanFitTranscodedString(source, destination))
        {
            bytesWritten = 0;
            return false;
        }

        var index = source.IndexOfAny(SearchCharacters);
        if (index == -1)
        {
            bytesWritten = Utf8NoBom.GetBytes(source, destination);
            return true;
        }

        var initialDestinationLength = destination.Length;

        while (index != -1)
        {
            if (index > 0)
            {
                var written = Utf8NoBom.GetBytes(source.Slice(0, index), destination);
                source = source.Slice(index);
                destination = destination.Slice(written);
            }

            var characterToReplace = source[0];
            source = source.Slice(1);

            var replacement = ByteReplacements.GetValueOrDefault(characterToReplace, []);
            if (replacement.Length > 0)
            {
                replacement.CopyTo(destination);
                destination = destination.Slice(replacement.Length);
            }

            index = source.IndexOfAny(SearchCharacters);
        }

        var finalWritten = Utf8NoBom.GetBytes(source, destination);
        bytesWritten = initialDestinationLength - destination.Length + finalWritten;
        return true;
    }
#endif
}
