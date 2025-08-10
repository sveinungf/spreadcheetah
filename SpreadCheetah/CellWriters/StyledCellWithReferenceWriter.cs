using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<StyledCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in StyledCell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.StyleId is null
            ? writer.TryWriteCellWithReference(cell.DataCell, DefaultStyling, State)
            : writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, State);
    }

    protected override bool TryWriteCell(in StyledCell cell, StyleId styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteCellWithReference(cell.DataCell, actualStyleId, State);
    }

    protected override bool WriteStartElement(in StyledCell cell, StyleId? styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        return StringCellValueWriterBase.WriteStartElementWithReference(actualStyleId, State);
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.DataCell.Value.StringOrPrimitive.StringValue, ref cellValueIndex);
    }
}
