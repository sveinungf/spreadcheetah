using SpreadCheetah.Styling;
using System;

namespace SpreadCheetah.Helpers
{
    internal static class FormulaCellHelper
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

        private static readonly int MaxCellElementLength =
            BeginStyledStringFormulaCell.Length
            + SpreadsheetConstants.StyleIdMaxDigits
            + EndStyleBeginFormula.Length
            + EndFormulaBeginCachedValue.Length
            + EndCachedValueEndCell.Length;

        public static bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = buffer.GetRemainingBuffer();

            // Try with approximate formula text and cached value lengths
            var cellValueLength = (formulaText.Length + cachedValue.Value.Length) * Utf8Helper.MaxBytePerChar;
            if (MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                buffer.Index += GetBytes(formulaText, cachedValue, styleId, buffer.GetNextSpan(), false);
                return true;
            }

            // Try with more accurate lengths
            cellValueLength = Utf8Helper.GetByteCount(formulaText) + Utf8Helper.GetByteCount(cachedValue.Value);
            bytesNeeded = MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                buffer.Index += GetBytes(formulaText, cachedValue, styleId, buffer.GetNextSpan(), false);
                return true;
            }

            return false;
        }

        public static int GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, Span<byte> bytes, bool assertSize)
        {
            int bytesWritten;

            if (styleId is null)
            {
                var cellStart = cachedValue.DataType switch
                {
                    CellDataType.InlineString => BeginStringFormulaCell,
                    CellDataType.Boolean => BeginBooleanFormulaCell,
                    _ => BeginNumberFormulaCell
                };

                bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            }
            else
            {
                var cellStart = cachedValue.DataType switch
                {
                    CellDataType.InlineString => BeginStyledStringFormulaCell,
                    CellDataType.Boolean => StyledCellHelper.BeginStyledBooleanCell,
                    _ => StyledCellHelper.BeginStyledNumberCell
                };

                bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
                bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
                bytesWritten += SpanHelper.GetBytes(EndStyleBeginFormula, bytes.Slice(bytesWritten));
            }

            bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), assertSize);

            if (string.IsNullOrEmpty(cachedValue.Value))
            {
                bytesWritten += SpanHelper.GetBytes(EndFormulaEndCell, bytes.Slice(bytesWritten));
                return bytesWritten;
            }

            bytesWritten += SpanHelper.GetBytes(EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cachedValue.Value, bytes.Slice(bytesWritten), assertSize);
            bytesWritten += SpanHelper.GetBytes(EndCachedValueEndCell, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        public static int GetStartElementBytes(CellDataType dataType, StyleId? styleId, Span<byte> bytes)
        {
            if (styleId is null)
            {
                var span = dataType switch
                {
                    CellDataType.InlineString => BeginStringFormulaCell,
                    CellDataType.Boolean => BeginBooleanFormulaCell,
                    _ => BeginNumberFormulaCell
                };

                return SpanHelper.GetBytes(span, bytes);
            }

            var cellStart = dataType switch
            {
                CellDataType.InlineString => BeginStyledStringFormulaCell,
                CellDataType.Boolean => StyledCellHelper.BeginStyledBooleanCell,
                _ => StyledCellHelper.BeginStyledNumberCell
            };

            var bytesWritten = SpanHelper.GetBytes(cellStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(EndStyleBeginFormula, bytes.Slice(bytesWritten));
            return bytesWritten;
        }
    }
}
