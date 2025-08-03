using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

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
        return CellValueWriter.GetWriter(cell.Type).WriteStartElementWithReference(styleId, State);
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
