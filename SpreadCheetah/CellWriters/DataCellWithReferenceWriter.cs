using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
     : BaseCellWriter<DataCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in DataCell cell)
    {
        return cell.Writer.TryWriteCellWithReference(cell, DefaultStyling, State);
    }

    protected override bool WriteStartElement(in DataCell cell)
    {
        return cell.Writer.WriteStartElementWithReference(State);
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
