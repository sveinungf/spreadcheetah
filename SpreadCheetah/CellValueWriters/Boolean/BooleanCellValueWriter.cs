using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal abstract class BooleanCellValueWriter : CellValueWriter
    {
        private static readonly int DataCellElementLength =
            DataCellHelper.TrueBooleanCell.Length;

        private static readonly int StyledCellElementLength =
            StyledCellHelper.BeginStyledBooleanCell.Length +
            SpreadsheetConstants.StyleIdMaxDigits +
            StyledCellHelper.EndStyleTrueBooleanValue.Length;

        private static readonly int FormulaCellElementLength =
            StyledCellHelper.BeginStyledBooleanCell.Length +
            SpreadsheetConstants.StyleIdMaxDigits +
            FormulaCellHelper.EndStyleBeginFormula.Length +
            FormulaCellHelper.EndFormulaTrueBooleanValue.Length;

        protected abstract ReadOnlySpan<byte> DataCellBytes();
        protected abstract ReadOnlySpan<byte> EndFormulaValueBytes();
        protected abstract ReadOnlySpan<byte> EndStyleValueBytes();

        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer) => GetBytes(buffer);
        public override bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(styleId, buffer);

        private bool GetBytes(SpreadsheetBuffer buffer)
        {
            buffer.Advance(SpanHelper.GetBytes(DataCellBytes(), buffer.GetNextSpan()));
            return true;
        }

        private bool GetBytes(StyleId styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(EndStyleValueBytes(), bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            int bytesWritten;

            if (styleId is null)
            {
                bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginBooleanFormulaCell, bytes);
            }
            else
            {
                bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, bytes);
                bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
                bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
            }

            bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(EndFormulaValueBytes(), bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            bytesNeeded = DataCellElementLength;
            var remaining = buffer.GetRemainingBuffer();
            return bytesNeeded <= remaining && GetBytes(cell, buffer);
        }

        public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            bytesNeeded = StyledCellElementLength;
            var remaining = buffer.GetRemainingBuffer();
            return bytesNeeded <= remaining && GetBytes(cell, styleId, buffer);
        }

        public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            // Try with approximate formula text length
            bytesNeeded = FormulaCellElementLength + formulaText.Length * Utf8Helper.MaxBytePerChar;
            var remaining = buffer.GetRemainingBuffer();
            if (bytesNeeded <= remaining)
                return GetBytes(formulaText, cachedValue, styleId, buffer);

            // Try with more accurate length
            bytesNeeded = FormulaCellElementLength + Utf8Helper.GetByteCount(formulaText);
            return bytesNeeded <= remaining && GetBytes(formulaText, cachedValue, styleId, buffer);
        }

        public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

        public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
        {
            if (cell.Formula is null)
                return true;

            var cellEnd = EndFormulaValueBytes();
            if (cellEnd.Length > buffer.GetRemainingBuffer())
                return false;

            buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetNextSpan()));
            return true;
        }

        public override bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
        {
            if (styleId is null)
            {
                buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginBooleanFormulaCell, buffer.GetNextSpan()));
                return true;
            }

            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, buffer.GetNextSpan());
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool WriteStartElement(SpreadsheetBuffer buffer) => GetBytes(buffer);

        public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(styleId, buffer);

        public override bool Equals(in CellValue value, in CellValue other) => true;
        public override int GetHashCodeFor(in CellValue value) => 0;
    }
}
