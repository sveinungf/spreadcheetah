using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWithReferenceAttributesWriter : BaseCellWriter<DataCell>
{
    public DataCellWithReferenceAttributesWriter(CellWriterState state, DefaultStyling? defaultStyling)
        : base(state, defaultStyling)
    {
    }

    protected override bool TryWriteCell(in DataCell cell)
    {
        return cell.Writer.TryWriteCellWithReferenceAttribute(cell, DefaultStyling, State);
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
