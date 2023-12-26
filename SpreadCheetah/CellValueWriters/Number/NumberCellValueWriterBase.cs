using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriterBase : CellValueWriter
{
    protected abstract int GetStyleId(StyleId styleId);

    protected static ReadOnlySpan<byte> BeginDataCell => "<c><v>"u8;
    protected static ReadOnlySpan<byte> EndStyleBeginValue => "\"><v>"u8; // TODO: Rename
    protected static ReadOnlySpan<byte> EndDefaultCell => "</v></c>"u8;

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginDataCell}");
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var style = GetStyleId(styleId);
        return buffer.TryWrite($"{StyledCellHelper.BeginStyledNumberCell}{style}{EndStyleBeginValue}");
    }

    public override bool WriteStartElementWithReference(CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndStyleBeginValue}");
    }

    public override bool WriteStartElementWithReference(StyleId styleId, CellWriterState state)
    {
        var style = GetStyleId(styleId);
        return state.Buffer.TryWrite($"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{EndStyleBeginValue}");
    }

    protected static bool WriteFormulaStartElement(int? styleId, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndStyleBeginFormula}")
            : buffer.TryWrite($"{FormulaCellHelper.BeginNumberFormulaCell}");
    }

    protected static bool WriteFormulaStartElementWithReference(int? styleId, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{FormulaCellHelper.EndStyleBeginFormula}")
            : state.Buffer.TryWrite($"{state}{FormulaCellHelper.EndStyleBeginFormula}");
    }

    public override bool CanWriteValuePieceByPiece(in DataCell cell) => true;

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{EndDefaultCell}");
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null
            ? TryWriteEndElement(buffer)
            : buffer.TryWrite($"{FormulaCellHelper.EndCachedValueEndCell}");
    }
}
