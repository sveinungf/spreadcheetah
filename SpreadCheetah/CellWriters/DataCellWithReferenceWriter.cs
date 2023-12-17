using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
     : BaseCellWriter<DataCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in DataCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool WriteStartElement(in DataCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool TryWriteEndElement(in DataCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
    {
        throw new NotImplementedException();
    }
}
