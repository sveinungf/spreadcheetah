using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.String
{
    internal sealed class NullStringCellValueWriter : BaseStringCellValueWriter
    {
        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
        {
            buffer.Advance(SpanHelper.GetBytes(DataCellHelper.NullStringCell, buffer.GetNextSpan()));
            return true;
        }

        public override bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
        {
            var bytes = buffer.GetNextSpan();
            var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleNullStringValue, bytes.Slice(bytesWritten));
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
                bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
                bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
                bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
            }

            bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaEndCell, bytes.Slice(bytesWritten));
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
    }
}
