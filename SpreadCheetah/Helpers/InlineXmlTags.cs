namespace SpreadCheetah.Helpers;

/// <summary>
/// Holds the XML tag fragments that vary depending on whether
/// <see cref="SpreadCheetahOptions.PreserveStringWhitespace"/> is enabled.
/// Each variant is resolved once into pre-baked byte arrays so that hot
/// write paths can read the spans directly without branching per cell.
/// </summary>
internal sealed class InlineXmlTags
{
    /// <summary>Cached instance used when whitespace preservation is disabled.</summary>
    public static readonly InlineXmlTags Default = new(preserveStringWhitespace: false);

    /// <summary>Cached instance used when whitespace preservation is enabled.</summary>
    public static readonly InlineXmlTags Preserve = new(preserveStringWhitespace: true);

    /// <summary>Returns <see cref="Preserve"/> when <paramref name="preserveStringWhitespace"/> is true, otherwise <see cref="Default"/>.</summary>
    public static InlineXmlTags For(bool preserveStringWhitespace) => preserveStringWhitespace ? Preserve : Default;

    private readonly byte[] _beginStringCell;
    private readonly byte[] _endReferenceBeginString;
    private readonly byte[] _endStyleBeginInlineString;
    private readonly byte[] _commentAfterRef;

    /// <summary>Opening tag for an unstyled inline string cell.</summary>
    public ReadOnlySpan<byte> BeginStringCell => _beginStringCell;

    /// <summary>Tag fragment that follows a cell reference attribute and opens the inline string content.</summary>
    public ReadOnlySpan<byte> EndReferenceBeginString => _endReferenceBeginString;

    /// <summary>Tag fragment that closes a style attribute and opens the inline string content.</summary>
    public ReadOnlySpan<byte> EndStyleBeginInlineString => _endStyleBeginInlineString;

    /// <summary>Tag fragment that follows a comment reference attribute and opens the comment text content.</summary>
    public ReadOnlySpan<byte> CommentAfterRef => _commentAfterRef;

    private InlineXmlTags(bool preserveStringWhitespace)
    {
        _beginStringCell = preserveStringWhitespace
            ? "<c t=\"inlineStr\"><is><t xml:space=\"preserve\">"u8.ToArray()
            : "<c t=\"inlineStr\"><is><t>"u8.ToArray();

        _endReferenceBeginString = preserveStringWhitespace
            ? "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">"u8.ToArray()
            : "\" t=\"inlineStr\"><is><t>"u8.ToArray();

        _endStyleBeginInlineString = preserveStringWhitespace
            ? "\"><is><t xml:space=\"preserve\">"u8.ToArray()
            : "\"><is><t>"u8.ToArray();

        _commentAfterRef = preserveStringWhitespace
            ? "\" authorId=\"0\"><text><r><t xml:space=\"preserve\">"u8.ToArray()
            : "\" authorId=\"0\"><text><r><t>"u8.ToArray();
    }
}
