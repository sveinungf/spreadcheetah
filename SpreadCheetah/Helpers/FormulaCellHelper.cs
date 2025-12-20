using SpreadCheetah.CellValueWriters;

namespace SpreadCheetah.Helpers;

internal static class FormulaCellHelper
{
    public static ReadOnlySpan<byte> EndQuoteBeginFormula => "\"><f>"u8;
    public static ReadOnlySpan<byte> BeginNumberFormulaCell => "<c><f>"u8;
    public static ReadOnlySpan<byte> EndFormulaBeginCachedValue => "</f><v>"u8;
    public static ReadOnlySpan<byte> EndCachedValueEndCell => "</v></c>"u8;

    public static bool FinishWritingFormulaCellValue(in Cell cell, string formulaText, ref int cellValueIndex, SpreadsheetBuffer buffer)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);

        // Write the formula
        if (cellValueIndex < formulaText.Length)
        {
            if (!buffer.WriteLongStringXmlEncoded(formulaText, ref cellValueIndex))
                return false;

            if (cell.DataCell.Type is CellWriterType.Null or CellWriterType.NullDateTime)
                return true;
        }

        // If there is a cached value, we need to write "[FORMULA]</f><v>[CACHEDVALUE]"
        var cachedValueStartIndex = formulaText.Length + 1;

        // Write the "</f><v>" part
        if (cellValueIndex < cachedValueStartIndex)
        {
            if (!buffer.TryWrite(EndFormulaBeginCachedValue))
                return false;

            cellValueIndex = cachedValueStartIndex;
        }

        // Write the cached value
        var cachedValueIndex = cellValueIndex - cachedValueStartIndex;
        var result = writer.WriteValuePieceByPiece(cell.DataCell, buffer, ref cachedValueIndex);
        cellValueIndex = cachedValueIndex + cachedValueStartIndex;
        return result;
    }
}
