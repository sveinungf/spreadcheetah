using System;

namespace SpreadCheetah.Helpers
{
    internal static class DataCellHelper
    {
        private static ReadOnlySpan<byte> NumberCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        private static ReadOnlySpan<byte> BooleanCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        public static ReadOnlySpan<byte> DefaultCellEnd => new[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        private static ReadOnlySpan<byte> StringCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i', (byte)'n', (byte)'e',
            (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };

        public static ReadOnlySpan<byte> StringCellEnd => new[]
        {
            (byte)'<', (byte)'/', (byte)'t', (byte)'>', (byte)'<', (byte)'/', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        private static readonly int MaxCellElementLength = StringCellStart.Length + StringCellEnd.Length;

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
                    cellStart = StringCellStart;
                    cellEnd = StringCellEnd;
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart;
                    cellEnd = DefaultCellEnd;
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart;
                    cellEnd = DefaultCellEnd;
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
            CellDataType.InlineString => SpanHelper.GetBytes(StringCellStart, bytes),
            CellDataType.Number => SpanHelper.GetBytes(NumberCellStart, bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(BooleanCellStart, bytes),
            _ => 0
        };

        public static bool TryWriteEndElement(in DataCell cell, SpreadsheetBuffer buffer)
        {
            var cellEnd = cell.DataType == CellDataType.InlineString
                ? StringCellEnd
                : DefaultCellEnd;

            if (cellEnd.Length > buffer.GetRemainingBuffer())
                return false;

            buffer.Index += SpanHelper.GetBytes(cellEnd, buffer.GetNextSpan());
            return true;
        }
    }
}
