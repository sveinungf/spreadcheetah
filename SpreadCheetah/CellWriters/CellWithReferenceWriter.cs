using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWithReferenceWriter : ICellWriter<Cell>
{
    public static CellWithReferenceWriter Instance { get; } = new();

    public bool TryWrite(in Cell cell, CellWriterState state)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell switch
        {
            { Formula: { } f } => writer.TryWriteCellWithReference(f.FormulaText, cell.DataCell, cell.StyleId, state.DefaultStyling, state),
            { StyleId: not null } => writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, state),
            _ => writer.TryWriteCellWithReference(cell.DataCell, state.DefaultStyling, state)
        };
    }

    public bool TryWrite(in Cell cell, StyleId styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.Formula is { } formula
            ? writer.TryWriteCellWithReference(formula.FormulaText, cell.DataCell, actualStyleId, state.DefaultStyling, state)
            : writer.TryWriteCellWithReference(cell.DataCell, actualStyleId, state);
    }

    public void WriteStartElement(in Cell cell, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;

        if (cell.Formula is not null)
        {
            var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
            var ok = writer.WriteFormulaStartElementWithReference(actualStyleId, state.DefaultStyling, state);
            Debug.Assert(ok);
            return;
        }

        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElementWithReference(actualStyleId, state);
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
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteEndElement(cell, state.Buffer);
    }
}