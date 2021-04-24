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

        // <c t="b"><v>
        private static ReadOnlySpan<byte> BeginBooleanCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // </v></c>
        public static ReadOnlySpan<byte> EndDefaultCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // <c t="inlineStr"><is><t>
        private static ReadOnlySpan<byte> BeginStringCell => new[]
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

        private static readonly int MaxCellElementLength = BeginStringCell.Length + EndStringCell.Length;

        public static bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            var cellValueLength = cell.Value.Length * Utf8Helper.MaxBytePerChar;
            if (MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                buffer.Index += GetBytes(cell, buffer.GetNextSpan(), false);
                return true;
            }

            // Try with a more accurate cell value length
            cellValueLength = Utf8Helper.GetByteCount(cell.Value);
            bytesNeeded = MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                buffer.Index += GetBytes(cell, buffer.GetNextSpan(), false);
                return true;
            }

            return false;
        }

        public static int GetBytes(in DataCell cell, Span<byte> bytes, bool assertSize)
        {
            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = BeginStringCell;
                    cellEnd = EndStringCell;
                    break;
                case CellDataType.Number:
                    cellStart = BeginNumberCell;
                    cellEnd = EndDefaultCell;
                    break;
                case CellDataType.Boolean:
                    cellStart = BeginBooleanCell;
                    cellEnd = EndDefaultCell;
                    break;
                default:
                    return 0;
            }

            var bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(cell.Value, bytes.Slice(bytesWritten), assertSize);
            bytesWritten += SpanHelper.GetBytes(cellEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        public static int GetStartElementBytes(CellDataType type, Span<byte> bytes) => type switch
        {
            CellDataType.InlineString => SpanHelper.GetBytes(BeginStringCell, bytes),
            CellDataType.Number => SpanHelper.GetBytes(BeginNumberCell, bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(BeginBooleanCell, bytes),
            _ => 0
        };

        public static bool TryWriteEndElement(in DataCell cell, SpreadsheetBuffer buffer)
        {
            var cellEnd = cell.DataType == CellDataType.InlineString
                ? EndStringCell
                : EndDefaultCell;

            if (cellEnd.Length > buffer.GetRemainingBuffer())
                return false;

            buffer.Index += SpanHelper.GetBytes(cellEnd, buffer.GetNextSpan());
            return true;
        }
    }
}
