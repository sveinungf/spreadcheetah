using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriter : ICellWriter<Cell>
{
    public static CellWriter Instance { get; } = new();

    public bool TryWrite(in Cell cell, CellWriterState state)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell switch
        {
            { Formula: { } formula } => writer.TryWriteCell(formula.FormulaText, cell.DataCell, cell.StyleId, state),
            { StyleId: not null } => writer.TryWriteCell(cell.DataCell, cell.StyleId, state.Buffer),
            _ => writer.TryWriteCell(cell.DataCell, state)
        };
    }

    public bool TryWrite(in Cell cell, StyleId styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.Formula is { } formula
            ? writer.TryWriteCell(formula.FormulaText, cell.DataCell, actualStyleId, state)
            : writer.TryWriteCell(cell.DataCell, actualStyleId, state.Buffer);
    }

    public void WriteStartElement(in Cell cell, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;

        if (cell.Formula is not null)
        {
            var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
            var ok = writer.WriteFormulaStartElement(actualStyleId, state);
            Debug.Assert(ok);
            return;
        }

        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElement(actualStyleId, state.Buffer);
        Debug.Assert(result);
    }

    public bool TryWriteValue(in Cell cell, ref int valueIndex, CellWriterState state)
    {
        return cell.Formula is { } formula
            ? FormulaCellHelper.FinishWritingFormulaCellValue(cell, formula.FormulaText, ref valueIndex, state.Buffer)
            : CellValueWriter.GetWriter(cell.DataCell.Type).WriteValuePieceByPiece(cell.DataCell, state.Buffer, ref valueIndex);
    }

    public bool TryWriteEndElement(in Cell cell, CellWriterState state)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteEndElement(cell, state.Buffer);
    }
}