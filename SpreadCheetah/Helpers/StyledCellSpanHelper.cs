using SpreadCheetah.Styling;
using System;

namespace SpreadCheetah.Helpers
{
    internal static class StyledCellSpanHelper
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
        private static ReadOnlySpan<byte> BeginStyledStringCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'i', (byte)'n', (byte)'l', (byte)'i',
            (byte)'n', (byte)'e', (byte)'S', (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // "><v>
        private static ReadOnlySpan<byte> EndStyleBeginValue => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // "><is><t>
        private static ReadOnlySpan<byte> EndStyleBeginInlineString => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'i', (byte)'s', (byte)'>', (byte)'<', (byte)'t', (byte)'>'
        };

        public static readonly int MaxCellEndElementLength = DataCellSpanHelper.StringCellEnd.Length;
        public static readonly int MaxCellElementLength = BeginStyledStringCell.Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + EndStyleBeginInlineString.Length
            + DataCellSpanHelper.StringCellEnd.Length;

        public static int GetBytes(StyledCell styledCell, Span<byte> bytes, bool assertSize)
        {
            return GetBytes(styledCell.DataCell, styledCell.StyleId, bytes, assertSize);
        }

        public static int GetBytes(DataCell cell, StyleId? styleId, Span<byte> bytes, bool assertSize)
        {
            if (styleId is null)
                return DataCellSpanHelper.GetBytes(cell, bytes, assertSize);

            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = BeginStyledStringCell;
                    cellStartEndTag = EndStyleBeginInlineString;
                    cellEnd = DataCellSpanHelper.StringCellEnd;
                    break;
                case CellDataType.Number:
                    cellStart = BeginStyledNumberCell;
                    cellStartEndTag = EndStyleBeginValue;
                    cellEnd = DataCellSpanHelper.DefaultCellEnd;
                    break;
                case CellDataType.Boolean:
                    cellStart = BeginStyledBooleanCell;
                    cellStartEndTag = EndStyleBeginValue;
                    cellEnd = DataCellSpanHelper.DefaultCellEnd;
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
            var cell = styledCell.DataCell;
            var styleId = styledCell.StyleId;
            if (styleId is null)
                return DataCellSpanHelper.GetStartElementBytes(cell.DataType, bytes);

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
