using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<Cell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in Cell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool WriteStartElement(in Cell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool TryWriteEndElement(in Cell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
    {
        throw new NotImplementedException();
    }
}
