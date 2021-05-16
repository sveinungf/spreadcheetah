using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters
{
    internal sealed class StringCellValueWriter : CellValueWriter
    {
        private static readonly int DataCellElementLength =
            DataCellHelper.BeginStringCell.Length +
            DataCellHelper.EndStringCell.Length;

        private static readonly int StyledCellElementLength =
            StyledCellHelper.BeginStyledStringCell.Length +
            SpreadsheetConstants.StyleIdMaxDigits +
            StyledCellHelper.EndStyleBeginInlineString.Length +
            DataCellHelper.EndStringCell.Length;

        private static readonly int FormulaCellElementLength =
            FormulaCellHelper.BeginStyledStringFormulaCell.Length +
            SpreadsheetConstants.StyleIdMaxDigits +
            FormulaCellHelper.EndStyleBeginFormula.Length +
            FormulaCellHelper.EndFormulaBeginCachedValue.Length +
            FormulaCellHelper.EndCachedValueEndCell.Length;

        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(DataCellHelper.BeginStringCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(cell.StringValue!, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndStringCell, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginInlineString, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cell.StringValue!, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndStringCell, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            int bytesWritten;

            if (styleId is null)
            {
                bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginStringFormulaCell, bytes);
            }
            else
            {
                bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginStyledStringFormulaCell, bytes);
                bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
                bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
            }

            bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
            bytesWritten += Utf8Helper.GetBytes(cachedValue.StringValue!, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndCachedValueEndCell, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            var remaining = buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            bytesNeeded = DataCellElementLength + cell.StringValue!.Length * Utf8Helper.MaxBytePerChar;
            if (bytesNeeded <= remaining)
                return GetBytes(cell, buffer);

            // Try with a more accurate cell value length
            bytesNeeded = DataCellElementLength + Utf8Helper.GetByteCount(cell.StringValue);
            return bytesNeeded <= remaining && GetBytes(cell, buffer);
        }

        public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            var remaining = buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            bytesNeeded = StyledCellElementLength + cell.StringValue!.Length * Utf8Helper.MaxBytePerChar;
            if (bytesNeeded <= remaining)
                return GetBytes(cell, styleId, buffer);

            // Try with a more accurate length
            bytesNeeded = StyledCellElementLength + Utf8Helper.GetByteCount(cell.StringValue);
            return bytesNeeded <= remaining && GetBytes(cell, styleId, buffer);
        }

        public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
        {
            var remaining = buffer.GetRemainingBuffer();

            // Try with an approximate cell value and formula text length
            bytesNeeded = FormulaCellElementLength + (formulaText.Length + cachedValue.StringValue!.Length) * Utf8Helper.MaxBytePerChar;
            if (bytesNeeded <= remaining)
                return GetBytes(formulaText, cachedValue, styleId, buffer);

            // Try with more accurate length
            bytesNeeded = FormulaCellElementLength + Utf8Helper.GetByteCount(formulaText) + Utf8Helper.GetByteCount(cachedValue.StringValue);
            return bytesNeeded <= remaining && GetBytes(formulaText, cachedValue, styleId, buffer);
        }

        public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
        {
            var cellEnd = DataCellHelper.EndStringCell;
            var bytes = buffer.GetNextSpan();
            if (cellEnd.Length >= bytes.Length)
                return false;

            buffer.Advance(SpanHelper.GetBytes(cellEnd, bytes));
            return true;
        }

        public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
        {
            if (cell.Formula is null)
                return TryWriteEndElement(buffer);

            var cellEnd = FormulaCellHelper.EndCachedValueEndCell;
            if (cellEnd.Length > buffer.GetRemainingBuffer())
                return false;

            buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetNextSpan()));
            return true;
        }

        public override bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
        {
            if (styleId is null)
            {
                buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginStringFormulaCell, buffer.GetNextSpan()));
                return true;
            }

            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool WriteStartElement(SpreadsheetBuffer buffer)
        {
            buffer.Advance(SpanHelper.GetBytes(DataCellHelper.BeginStringCell, buffer.GetNextSpan()));
            return true;
        }

        public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginInlineString, bytes.Slice(bytesWritten));
            buffer.Advance(bytesWritten);
            return true;
        }

        public override bool Equals(in CellValue value, in CellValue other) => true;
        public override int GetHashCodeFor(in CellValue value) => 0;
    }
}
