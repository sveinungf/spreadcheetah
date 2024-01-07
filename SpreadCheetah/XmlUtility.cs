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
#else
    private static readonly SearchValues<char> SearchCharacters;
    private static readonly FrozenDictionary<char, string> Replacements;

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
#endif
}
