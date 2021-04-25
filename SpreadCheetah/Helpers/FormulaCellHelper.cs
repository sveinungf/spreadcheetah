using System;

namespace SpreadCheetah.Helpers
{
    internal static class FormulaCellHelper
    {
        // <c t="str" s="
        public static ReadOnlySpan<byte> BeginStyledStringFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'s',
            (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // "><f>
        public static ReadOnlySpan<byte> EndStyleBeginFormula => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c><f>
        public static ReadOnlySpan<byte> BeginNumberFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c t="b"><f>
        public static ReadOnlySpan<byte> BeginBooleanFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b',
            (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c t="str"><f>
        public static ReadOnlySpan<byte> BeginStringFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'s',
            (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // </f><v>
        public static ReadOnlySpan<byte> EndFormulaBeginCachedValue => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // </v></c>
        public static ReadOnlySpan<byte> EndCachedValueEndCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // </f></c>
        public static ReadOnlySpan<byte> EndFormulaEndCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // </f><v>0</v></c>
        public static ReadOnlySpan<byte> EndFormulaFalseBooleanValue => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'v', (byte)'>', (byte)'0',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // </f><v>1</v></c>
        public static ReadOnlySpan<byte> EndFormulaTrueBooleanValue => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'v', (byte)'>', (byte)'1',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };
    }
}
