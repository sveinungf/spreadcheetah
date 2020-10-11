using System;

namespace SpreadCheetah.Helpers
{
    internal static class CellSpanHelper
    {
        private static ReadOnlySpan<byte> NumberCellStart => new byte[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        private static ReadOnlySpan<byte> BooleanCellStart => new byte[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        private static ReadOnlySpan<byte> DefaultCellEnd => new byte[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        private static ReadOnlySpan<byte> StringCellStart => new byte[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i', (byte)'n', (byte)'e',
            (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };

        private static ReadOnlySpan<byte> StringCellEnd => new byte[]
        {
            (byte)'<', (byte)'/', (byte)'t', (byte)'>', (byte)'<', (byte)'/', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        public static readonly int MaxCellEndElementLength = StringCellEnd.Length;
        public static readonly int MaxCellElementLength = StringCellStart.Length + MaxCellEndElementLength;

        public static int GetBytes(in Cell cell, Span<byte> bytes)
        {
            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    StringCellStart.CopyTo(bytes);
                    var a = Utf8Helper.GetBytes(cell.Value, bytes.Slice(StringCellStart.Length), false);
                    StringCellEnd.CopyTo(bytes.Slice(StringCellStart.Length + a));
                    return StringCellStart.Length + StringCellEnd.Length + a;
                case CellDataType.Number:
                    NumberCellStart.CopyTo(bytes);
                    var b = Utf8Helper.GetBytes(cell.Value, bytes.Slice(NumberCellStart.Length));
                    DefaultCellEnd.CopyTo(bytes.Slice(NumberCellStart.Length + b));
                    return NumberCellStart.Length + DefaultCellEnd.Length + b;
                case CellDataType.Boolean:
                    BooleanCellStart.CopyTo(bytes);
                    var c = Utf8Helper.GetBytes(cell.Value, bytes.Slice(BooleanCellStart.Length));
                    DefaultCellEnd.CopyTo(bytes.Slice(BooleanCellStart.Length + c));
                    return BooleanCellStart.Length + DefaultCellEnd.Length + c;
                default:
                    return 0;
            }
        }

        public static int GetStartElementBytes(CellDataType type, Span<byte> bytes) => type switch
        {
            CellDataType.InlineString => SpanHelper.GetBytes(StringCellStart, bytes),
            CellDataType.Number => SpanHelper.GetBytes(NumberCellStart, bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(BooleanCellStart, bytes),
            _ => 0
        };

        public static int GetEndElementBytes(CellDataType type, Span<byte> bytes) => type switch
        {
            CellDataType.InlineString => SpanHelper.GetBytes(StringCellEnd, bytes),
            CellDataType.Number => SpanHelper.GetBytes(DefaultCellEnd, bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(DefaultCellEnd, bytes),
            _ => 0
        };
    }
}
