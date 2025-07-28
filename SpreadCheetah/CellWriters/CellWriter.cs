using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<Cell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in Cell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell switch
        {
            { Formula: { } formula } => writer.TryWriteCell(formula.FormulaText, cell.DataCell, cell.StyleId, DefaultStyling, Buffer),
            { StyleId: not null } => writer.TryWriteCell(cell.DataCell, cell.StyleId, Buffer),
            _ => writer.TryWriteCell(cell.DataCell, DefaultStyling, Buffer)
        };
    }

    protected override bool TryWriteCell(in Cell cell, StyleId styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.Formula is { } formula
            ? writer.TryWriteCell(formula.FormulaText, cell.DataCell, actualStyleId, DefaultStyling, Buffer)
            : writer.TryWriteCell(cell.DataCell, actualStyleId, Buffer);
    }

    protected override bool WriteStartElement(in Cell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell switch
        {
            { Formula: not null } => writer.WriteFormulaStartElement(cell.StyleId, DefaultStyling, Buffer),
            { StyleId: not null } => writer.WriteStartElement(cell.StyleId, Buffer),
            _ => writer.WriteStartElement(Buffer)
        };
    }

    protected override bool TryWriteEndElement(in Cell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteEndElement(cell, Buffer);
    }

    protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
    {
        return cell.Formula is { } formula
            ? FinishWritingFormulaCellValue(cell, formula.FormulaText, ref cellValueIndex)
            : CellValueWriter.GetWriter(cell.DataCell.Type).WriteValuePieceByPiece(cell.DataCell, Buffer, ref cellValueIndex);
    }
}
