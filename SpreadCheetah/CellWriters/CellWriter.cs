using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<Cell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in Cell cell) => cell switch
    {
        { Formula: { } formula } => cell.DataCell.Writer.TryWriteCell(formula.FormulaText, cell.DataCell, cell.StyleId, DefaultStyling, Buffer),
        { StyleId: not null } => cell.DataCell.Writer.TryWriteCell(cell.DataCell, cell.StyleId, Buffer),
        _ => cell.DataCell.Writer.TryWriteCell(cell.DataCell, DefaultStyling, Buffer)
    };

    protected override bool WriteStartElement(in Cell cell) => cell switch
    {
        { Formula: not null } => cell.DataCell.Writer.WriteFormulaStartElement(cell.StyleId, DefaultStyling, Buffer),
        { StyleId: not null } => cell.DataCell.Writer.WriteStartElement(cell.StyleId, Buffer),
        _ => cell.DataCell.Writer.WriteStartElement(Buffer)
    };

    protected override bool TryWriteEndElement(in Cell cell)
    {
        return cell.DataCell.Writer.TryWriteEndElement(cell, Buffer);
    }

    protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
    {
        return cell.Formula is { } formula
            ? FinishWritingFormulaCellValue(cell, formula.FormulaText, ref cellValueIndex)
            : cell.DataCell.Writer.WriteValuePieceByPiece(cell.DataCell, Buffer, ref cellValueIndex);
    }
}
