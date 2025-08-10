using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
     : BaseCellWriter<DataCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in DataCell cell)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCellWithReference(cell, DefaultStyling, State);
    }

    protected override bool TryWriteCell(in DataCell cell, StyleId styleId)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCellWithReference(cell, styleId, State);
    }

    protected override bool WriteStartElement(in DataCell cell, StyleId? styleId)
    {
        Debug.Assert(CellValueWriter.GetWriter(cell.Type) is StringCellValueWriterBase);
        return StringCellValueWriterBase.WriteStartElementWithReference(styleId, State);
    }

    protected override bool TryWriteEndElement(in DataCell cell)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.Value.StringOrPrimitive.StringValue, ref cellValueIndex);
    }
}
