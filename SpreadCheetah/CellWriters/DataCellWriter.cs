namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWriter : BaseCellWriter<DataCell>
{
    public DataCellWriter(SpreadsheetBuffer buffer) : base(buffer)
    {
    }

    protected override bool TryWriteCell(in DataCell cell)
    {
        return cell.Writer.TryWriteCell(cell, Buffer);
    }

    protected override bool WriteStartElement(in DataCell cell)
    {
        return cell.Writer.WriteStartElement(Buffer);
    }

    protected override bool TryWriteEndElement(in DataCell cell)
    {
        return cell.Writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
    {
        return FinishWritingCellValue(cell.StringValue!, ref cellValueIndex);
    }
}
