using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class DoubleCellValueWriter : NumberCellValueWriter
{
    protected override bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten)
    {
        return Utf8Formatter.TryFormat(cell.NumberValue.DoubleValue, destination, out bytesWritten);
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginDataCell}{cell.NumberValue.DoubleValue}{EndDefaultCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(
            $"{StyledCellHelper.BeginStyledNumberCell}{styleId.Id}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{styleId.Id}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
            $"{EndDefaultCell}");
    }
}
