using System;

namespace SpreadCheetah.Helpers
{
    internal static partial class StyledCellSpanHelper
    {
        [StringLiteral.Utf8("<c s=\"")]
        public static partial ReadOnlySpan<byte> NumberCellStart();

        [StringLiteral.Utf8("\"><v>")]
        private static partial ReadOnlySpan<byte> NumberCellStartEndTag();

        [StringLiteral.Utf8("<c t=\"b\" s=\"")]
        private static partial ReadOnlySpan<byte> BooleanCellStart();

        [StringLiteral.Utf8("\"><v>")]
        private static partial ReadOnlySpan<byte> BooleanCellStartEndTag();

        [StringLiteral.Utf8("<c t=\"inlineStr\" s=\"")]
        private static partial ReadOnlySpan<byte> StringCellStart();

        [StringLiteral.Utf8("\"><is><t>")]
        private static partial ReadOnlySpan<byte> StringCellStartEndTag();

        public static readonly int MaxCellEndElementLength = CellSpanHelper.StringCellEnd().Length;
        public static readonly int MaxCellElementLength = StringCellStart().Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + StringCellStartEndTag().Length
            + CellSpanHelper.StringCellEnd().Length;

        public static int GetBytes(StyledCell styledCell, Span<byte> bytes, bool assertSize)
        {
            var cell = styledCell.DataCell;
            var styleId = styledCell.StyleId;
            if (styleId is null)
                return CellSpanHelper.GetBytes(cell, bytes, assertSize);

            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = StringCellStart();
                    cellStartEndTag = StringCellStartEndTag();
                    cellEnd = CellSpanHelper.StringCellEnd();
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart();
                    cellStartEndTag = NumberCellStartEndTag();
                    cellEnd = CellSpanHelper.DefaultCellEnd();
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart();
                    cellStartEndTag = BooleanCellStartEndTag();
                    cellEnd = CellSpanHelper.DefaultCellEnd();
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
                return CellSpanHelper.GetStartElementBytes(cell.DataType, bytes);

            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellStartEndTag;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = StringCellStart();
                    cellStartEndTag = StringCellStartEndTag();
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart();
                    cellStartEndTag = NumberCellStartEndTag();
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart();
                    cellStartEndTag = BooleanCellStartEndTag();
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
