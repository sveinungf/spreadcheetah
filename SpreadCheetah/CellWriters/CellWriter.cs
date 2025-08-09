using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

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

    protected override bool WriteStartElement(in Cell cell, StyleId? styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;

        if (cell.Formula is not null)
        {
            var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
            return writer.WriteFormulaStartElement(actualStyleId, DefaultStyling, Buffer);
        }

        Debug.Assert(cell.DataCell.Type is CellWriterType.String or CellWriterType.ReadOnlyMemoryOfChar);
        return StringCellValueWriterBase.WriteStartElement(actualStyleId, Buffer);
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
