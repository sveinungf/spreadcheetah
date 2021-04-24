using SpreadCheetah.Styling;
using System;

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

        // "><v></v></c>
        public static ReadOnlySpan<byte> EndStyleNullValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>',
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
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

        // "><is><t></t></is></c>
        public static ReadOnlySpan<byte> EndStyleNullStringValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>',
            (byte)'<', (byte)'/', (byte)'t', (byte)'>', (byte)'<', (byte)'/', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        private static readonly int MaxCellElementLength =
            BeginStyledStringCell.Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + EndStyleBeginInlineString.Length
            + DataCellHelper.EndStringCell.Length;

        public static bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            var cellValueLength = cell.Value.Length * Utf8Helper.MaxBytePerChar;
            if (MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                buffer.Index += GetBytes(cell, styleId, buffer.GetNextSpan(), false);
                return true;
            }

            // Try with a more accurate cell value length
            cellValueLength = Utf8Helper.GetByteCount(cell.Value);
            bytesNeeded = MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                buffer.Index += GetBytes(cell, styleId, buffer.GetNextSpan(), false);
                return true;
            }

            return false;
        }

        public static int GetBytes(DataCell cell, StyleId styleId, Span<byte> bytes, bool assertSize)
        {
            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = BeginStyledStringCell;
                    cellStartEndTag = EndStyleBeginInlineString;
                    cellEnd = DataCellHelper.EndStringCell;
                    break;
                case CellDataType.Number:
                    cellStart = BeginStyledNumberCell;
                    cellStartEndTag = EndStyleBeginValue;
                    cellEnd = DataCellHelper.EndDefaultCell;
                    break;
                case CellDataType.Boolean:
                    cellStart = BeginStyledBooleanCell;
                    cellStartEndTag = EndStyleBeginValue;
                    cellEnd = DataCellHelper.EndDefaultCell;
                    break;
                default:
                    return 0;
            }

            var bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(cellStartEndTag, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cell.Value, bytes.Slice(bytesWritten), assertSize);
            bytesWritten += SpanHelper.GetBytes(cellEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        public static int GetStartElementBytes(DataCell cell, StyleId styleId, Span<byte> bytes)
        {
            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = BeginStyledStringCell;
                    cellStartEndTag = EndStyleBeginInlineString;
                    break;
                case CellDataType.Number:
                    cellStart = BeginStyledNumberCell;
                    cellStartEndTag = EndStyleBeginValue;
                    break;
                case CellDataType.Boolean:
                    cellStart = BeginStyledBooleanCell;
                    cellStartEndTag = EndStyleBeginValue;
                    break;
                default:
                    return 0;
            }

            var bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(cellStartEndTag, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
