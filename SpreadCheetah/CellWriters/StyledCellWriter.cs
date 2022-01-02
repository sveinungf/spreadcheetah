namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWriter : BaseCellWriter<StyledCell>
{
    public StyledCellWriter(SpreadsheetBuffer buffer) : base(buffer)
    {
    }

    protected override bool TryWriteCell(in StyledCell cell)
    {
        return cell.StyleId is null
            ? cell.DataCell.Writer.TryWriteCell(cell.DataCell, Buffer)
            : cell.DataCell.Writer.TryWriteCell(cell.DataCell, cell.StyleId, Buffer);
    }

    protected override bool WriteStartElement(in StyledCell cell)
    {
        return cell.StyleId is null
            ? cell.DataCell.Writer.WriteStartElement(Buffer)
            : cell.DataCell.Writer.WriteStartElement(cell.StyleId, Buffer);
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        return cell.DataCell.Writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        return FinishWritingCellValue(cell.DataCell.StringValue!, ref cellValueIndex);
    }
}
