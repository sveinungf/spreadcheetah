using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<DataCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in DataCell cell)
    {
        return cell.Writer.TryWriteCell(cell, DefaultStyling, State);
    }

    protected override bool WriteStartElement(in DataCell cell)
    {
        return cell.Writer.WriteStartElement(State);
    }

    protected override bool TryWriteEndElement(in DataCell cell)
    {
        return cell.Writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.StringValue, ref cellValueIndex);
    }
}
