using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<DataCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in DataCell cell)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCell(cell, DefaultStyling, Buffer);
    }

    protected override bool WriteStartElement(in DataCell cell)
    {
        return CellValueWriter.GetWriter(cell.Type).WriteStartElement(Buffer);
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
