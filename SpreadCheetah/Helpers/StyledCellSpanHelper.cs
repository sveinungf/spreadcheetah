using System;

namespace SpreadCheetah.Helpers
{
    internal static class StyledCellSpanHelper
    {
        private static ReadOnlySpan<byte> NumberCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        private static ReadOnlySpan<byte> NumberCellStartEndTag => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        private static ReadOnlySpan<byte> BooleanCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        private static ReadOnlySpan<byte> BooleanCellStartEndTag => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        private static ReadOnlySpan<byte> StringCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i', (byte)'n', (byte)'e',
            (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        private static ReadOnlySpan<byte> StringCellStartEndTag => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };

        public static readonly int MaxCellEndElementLength = CellSpanHelper.StringCellEnd.Length;
        public static readonly int MaxCellElementLength = StringCellStart.Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + StringCellStartEndTag.Length
            + CellSpanHelper.StringCellEnd.Length;

        public static int GetBytes(StyledCell styledCell, Span<byte> bytes)
        {
            var cell = styledCell.Cell;
            var styleId = styledCell.StyleId;
            if (styleId is null)
                return CellSpanHelper.GetBytes(cell, bytes);

            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = StringCellStart;
                    cellStartEndTag = StringCellStartEndTag;
                    cellEnd = CellSpanHelper.StringCellEnd;
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart;
                    cellStartEndTag = NumberCellStartEndTag;
                    cellEnd = CellSpanHelper.DefaultCellEnd;
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart;
                    cellStartEndTag = BooleanCellStartEndTag;
                    cellEnd = CellSpanHelper.DefaultCellEnd;
                    break;
                default:
                    return 0;
            }

            var bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(cellStartEndTag, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cell.Value, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(cellEnd, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        public static int GetStartElementBytes(StyledCell styledCell, Span<byte> bytes)
        {
            var cell = styledCell.Cell;
            var styleId = styledCell.StyleId;
            if (styleId is null)
                return CellSpanHelper.GetStartElementBytes(cell.DataType, bytes);

            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = StringCellStart;
                    cellStartEndTag = StringCellStartEndTag;
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart;
                    cellStartEndTag = NumberCellStartEndTag;
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart;
                    cellStartEndTag = BooleanCellStartEndTag;
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
