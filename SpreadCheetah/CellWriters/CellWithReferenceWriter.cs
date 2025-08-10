using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<Cell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in Cell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell switch
        {
            { Formula: { } f } => writer.TryWriteCellWithReference(f.FormulaText, cell.DataCell, cell.StyleId, DefaultStyling, State),
            { StyleId: not null } => writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, State),
            _ => writer.TryWriteCellWithReference(cell.DataCell, DefaultStyling, State)
        };
    }

    protected override bool TryWriteCell(in Cell cell, StyleId styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.Formula is { } formula
            ? writer.TryWriteCellWithReference(formula.FormulaText, cell.DataCell, actualStyleId, DefaultStyling, State)
            : writer.TryWriteCellWithReference(cell.DataCell, actualStyleId, State);
    }

    protected override bool WriteStartElement(in Cell cell, StyleId? styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;

        if (cell.Formula is not null)
        {
            var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
            return writer.WriteFormulaStartElementWithReference(actualStyleId, DefaultStyling, State);
        }

        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        return StringCellValueWriterBase.WriteStartElementWithReference(actualStyleId, State);
    }

    protected override bool TryWriteEndElement(in Cell cell)
    {
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteEndElement(cell, Buffer);
    }

    protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
    {
        return cell.Formula is { } formula
            ? FinishWritingFormulaCellValue(cell, formula.FormulaText, ref cellValueIndex)
            : CellValueWriter.GetWriter(cell.DataCell.Type).WriteValuePieceByPiece(cell.DataCell, Buffer, ref cellValueIndex);
    }
}
