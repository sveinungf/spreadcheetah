using System;

namespace SpreadCheetah.Helpers
{
    internal static class DataCellHelper
    {
        // <c><v>
        public static ReadOnlySpan<byte> BeginNumberCell => new[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // </v></c>
        public static ReadOnlySpan<byte> EndDefaultCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="inlineStr"><is><t>
        public static ReadOnlySpan<byte> BeginStringCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i', (byte)'n', (byte)'e',
            (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };

        // </t></is></c>
        public static ReadOnlySpan<byte> EndStringCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'t', (byte)'>', (byte)'<', (byte)'/', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c><v></v></c>
        public static ReadOnlySpan<byte> NullNumberCell => new[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="b"><v></v></c>
        public static ReadOnlySpan<byte> NullBooleanCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="inlineStr"><is><t></t></is></c>
        public static ReadOnlySpan<byte> NullStringCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i', (byte)'n', (byte)'e',
            (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>',
            (byte)'<', (byte)'/', (byte)'t', (byte)'>', (byte)'<', (byte)'/', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="b"><v>0</v></c>
        public static ReadOnlySpan<byte> FalseBooleanCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>',
            (byte)'0', (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="b"><v>1</v></c>
        public static ReadOnlySpan<byte> TrueBooleanCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>',
            (byte)'1', (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };
    }
}
