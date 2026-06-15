namespace SpreadCheetah.Helpers;

internal sealed class InlineXmlTags
{
    public static readonly InlineXmlTags Default = new(preserveStringWhitespace: false);
    public static readonly InlineXmlTags Preserve = new(preserveStringWhitespace: true);

    public static InlineXmlTags For(bool preserveStringWhitespace) => preserveStringWhitespace ? Preserve : Default;

    private readonly byte[] _beginStringCell;
    private readonly byte[] _endReferenceBeginString;
    private readonly byte[] _endStyleBeginInlineString;
    private readonly byte[] _commentAfterRef;

    public ReadOnlySpan<byte> BeginStringCell => _beginStringCell;
    public ReadOnlySpan<byte> EndReferenceBeginString => _endReferenceBeginString;
    public ReadOnlySpan<byte> EndStyleBeginInlineString => _endStyleBeginInlineString;
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
