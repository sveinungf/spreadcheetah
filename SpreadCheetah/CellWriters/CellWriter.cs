using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriter : BaseCellWriter<Cell>
{
    public CellWriter(CellWriterState state, DefaultStyling? defaultStyling)
        : base(state, defaultStyling)
    {
    }

    protected override bool TryWriteCell(in Cell cell) => cell switch
    {
        { Formula: not null } => cell.DataCell.Writer.TryWriteCell(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, DefaultStyling, State),
        { StyleId: not null } => cell.DataCell.Writer.TryWriteCell(cell.DataCell, cell.StyleId, State),
        _ => cell.DataCell.Writer.TryWriteCell(cell.DataCell, DefaultStyling, State)
    };

    protected override bool WriteStartElement(in Cell cell) => cell switch
    {
        { Formula: not null } => cell.DataCell.Writer.WriteFormulaStartElement(cell.StyleId, DefaultStyling, Buffer),
        { StyleId: not null } => cell.DataCell.Writer.WriteStartElement(cell.StyleId, Buffer),
        _ => cell.DataCell.Writer.WriteStartElement(State)
    };

    protected override bool TryWriteEndElement(in Cell cell)
    {
        return cell.DataCell.Writer.TryWriteEndElement(cell, Buffer);
    }

    protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
    {
        if (cell.Formula is null)
            return cell.DataCell.Writer.WriteValuePieceByPiece(cell.DataCell, Buffer, ref cellValueIndex);

        var formulaText = cell.Formula.Value.FormulaText;

        // Write the formula
        if (cellValueIndex < formulaText.Length)
        {
            if (!Buffer.WriteLongString(formulaText, ref cellValueIndex)) return false;

            // Finish if there is no cached value to write piece by piece
            if (!cell.DataCell.Writer.CanWriteValuePieceByPiece(cell.DataCell)) return true;
        }

        // If there is a cached value, we need to write "[FORMULA]</f><v>[CACHEDVALUE]"
        var cachedValueStartIndex = formulaText.Length + 1;

        // Write the "</f><v>" part
        if (cellValueIndex < cachedValueStartIndex)
        {
            var separator = FormulaCellHelper.EndFormulaBeginCachedValue;
            if (separator.Length > Buffer.FreeCapacity) return false;
            Buffer.Advance(SpanHelper.GetBytes(separator, Buffer.GetSpan()));
            cellValueIndex = cachedValueStartIndex;
        }

        // Write the cached value
        var cachedValueIndex = cellValueIndex - cachedValueStartIndex;
        var result = cell.DataCell.Writer.WriteValuePieceByPiece(cell.DataCell, Buffer, ref cachedValueIndex);
        cellValueIndex = cachedValueIndex + cachedValueStartIndex;
        return result;
    }
}
