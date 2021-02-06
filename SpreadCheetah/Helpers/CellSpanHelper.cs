using System;

namespace SpreadCheetah.Helpers
{
    internal static partial class CellSpanHelper
    {
        [StringLiteral.Utf8("<c><v>")]
        private static partial ReadOnlySpan<byte> NumberCellStart();

        [StringLiteral.Utf8("<c t=\"b\"><v>")]
        private static partial ReadOnlySpan<byte> BooleanCellStart();

        [StringLiteral.Utf8("</v></c>")]
        public static partial ReadOnlySpan<byte> DefaultCellEnd();

        [StringLiteral.Utf8("<c t=\"inlineStr\"><is><t>")]
        private static partial ReadOnlySpan<byte> StringCellStart();

        [StringLiteral.Utf8("</t></is></c>")]
        public static partial ReadOnlySpan<byte> StringCellEnd();

        public static readonly int MaxCellEndElementLength = StringCellEnd().Length;
        public static readonly int MaxCellElementLength = StringCellStart().Length + MaxCellEndElementLength;

        public static int GetBytes(in DataCell cell, Span<byte> bytes, bool assertSize)
        {
            ReadOnlySpan<byte> cellStart;
            ReadOnlySpan<byte> cellEnd;

            switch (cell.DataType)
            {
                case CellDataType.InlineString:
                    cellStart = StringCellStart();
                    cellEnd = StringCellEnd();
                    break;
                case CellDataType.Number:
                    cellStart = NumberCellStart();
                    cellEnd = DefaultCellEnd();
                    break;
                case CellDataType.Boolean:
                    cellStart = BooleanCellStart();
                    cellEnd = DefaultCellEnd();
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
            CellDataType.InlineString => SpanHelper.GetBytes(StringCellStart(), bytes),
            CellDataType.Number => SpanHelper.GetBytes(NumberCellStart(), bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(BooleanCellStart(), bytes),
            _ => 0
        };

        public static int GetEndElementBytes(CellDataType type, Span<byte> bytes) => type switch
        {
            CellDataType.InlineString => SpanHelper.GetBytes(StringCellEnd(), bytes),
            CellDataType.Number => SpanHelper.GetBytes(DefaultCellEnd(), bytes),
            CellDataType.Boolean => SpanHelper.GetBytes(DefaultCellEnd(), bytes),
            _ => 0
        };
    }
}
