namespace SpreadCheetah.Helpers
{
    internal static class StyledCellHelper
    {
        // <c s="
        public static ReadOnlySpan<byte> BeginStyledNumberCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // <c t="b" s="
        public static ReadOnlySpan<byte> BeginStyledBooleanCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"',
            (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // <c t="inlineStr" s="
        public static ReadOnlySpan<byte> BeginStyledStringCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i',
            (byte)'n', (byte)'e', (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // "><v>
        public static ReadOnlySpan<byte> EndStyleBeginValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // "/>
        public static ReadOnlySpan<byte> EndStyleNullValue => new[]
        {
            (byte)'"', (byte)'/', (byte)'>'
        };

        // "><v>0</v></c>
        public static ReadOnlySpan<byte> EndStyleFalseBooleanValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>', (byte)'0',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // "><v>1</v></c>
        public static ReadOnlySpan<byte> EndStyleTrueBooleanValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>', (byte)'1',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // "><is><t>
        public static ReadOnlySpan<byte> EndStyleBeginInlineString => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };
    }
}
