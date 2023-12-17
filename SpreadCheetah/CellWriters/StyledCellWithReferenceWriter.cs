using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<StyledCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in StyledCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool WriteStartElement(in StyledCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        throw new NotImplementedException();
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        throw new NotImplementedException();
    }
}
