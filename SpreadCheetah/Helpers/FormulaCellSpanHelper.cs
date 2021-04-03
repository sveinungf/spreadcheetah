using System;

namespace SpreadCheetah.Helpers
{
    internal static class FormulaCellSpanHelper
    {
        // <c t="str" s="
        private static ReadOnlySpan<byte> BeginStyledStringFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'s',
            (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // "><f>
        private static ReadOnlySpan<byte> EndStyleBeginFormula => new[]
        {
            (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c><f>
        private static ReadOnlySpan<byte> BeginNumberFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c t="b"><f>
        private static ReadOnlySpan<byte> BeginBooleanFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'b',
            (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // <c t="str"><f>
        private static ReadOnlySpan<byte> BeginStringFormulaCell => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'s',
            (byte)'t', (byte)'r', (byte)'"', (byte)'>', (byte)'<', (byte)'f', (byte)'>'
        };

        // </f><v>
        private static ReadOnlySpan<byte> EndFormulaBeginCachedValue => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
        };

        // </v></c>
        private static ReadOnlySpan<byte> EndCachedValueEndCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'v', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // TODO: Try with skipping the <v> element when no cached value.
        // </f></c>
        private static ReadOnlySpan<byte> EndFormulaEndCell => new[]
        {
            (byte)'<', (byte)'/', (byte)'f', (byte)'>', (byte)'<', (byte)'/', (byte)'c', (byte)'>'
        };

        // TODO: Is this the longest? What about inlineStr?
        public static readonly int MaxCellElementLength =
            BeginStyledStringFormulaCell.Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + EndStyleBeginFormula.Length
            + EndFormulaBeginCachedValue.Length
            + EndCachedValueEndCell.Length;
        //<c t="str" s="1">
        //	<f>UPPER(B1)</f>
        //	<v>TEST</v>
        //</c>

        public static int GetBytes(Cell cell, Span<byte> bytes, bool assertSize)
        {
            if (cell.Formula is null)
                return StyledCellSpanHelper.GetBytes(cell.DataCell, cell.StyleId, bytes, assertSize);

            int bytesWritten;

            if (cell.StyleId is null)
            {
                var cellStart = cell.DataCell.DataType switch
                {
                    CellDataType.InlineString => BeginStringFormulaCell,
                    CellDataType.Boolean => BeginBooleanFormulaCell,
                    _ => BeginNumberFormulaCell
                };

                bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            }
            else
            {
                var cellStart = cell.DataCell.DataType switch
                {
                    CellDataType.InlineString => BeginStyledStringFormulaCell,
                    CellDataType.Boolean => StyledCellSpanHelper.BeginStyledBooleanCell,
                    _ => StyledCellSpanHelper.BeginStyledNumberCell
                };

                bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
                bytesWritten += Utf8Helper.GetBytes(cell.StyleId.Id, bytes.Slice(bytesWritten));
                bytesWritten += SpanHelper.GetBytes(EndStyleBeginFormula, bytes.Slice(bytesWritten));
            }

            bytesWritten += Utf8Helper.GetBytes(cell.Formula.Value.FormulaText, bytes.Slice(bytesWritten));

            if (string.IsNullOrEmpty(cell.DataCell.Value))
            {
                bytesWritten += SpanHelper.GetBytes(EndFormulaEndCell, bytes.Slice(bytesWritten));
                return bytesWritten;
            }

            bytesWritten += SpanHelper.GetBytes(EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cell.DataCell.Value, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(EndCachedValueEndCell, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
