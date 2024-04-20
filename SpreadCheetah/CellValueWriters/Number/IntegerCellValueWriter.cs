using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class IntegerCellValueWriter : NumberCellValueWriter
{
    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginDataCell}{cell.Value.StringOrPrimitive.PrimitiveValue.IntValue}{EndDefaultCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(
            $"{StyledCellHelper.BeginStyledNumberCell}{styleId.Id}{EndQuoteBeginValue}" +
            $"{cell.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{StyledCellHelper.BeginStyledNumberCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return buffer.TryWrite(
            $"{FormulaCellHelper.BeginNumberFormulaCell}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{EndQuoteBeginValue}" +
            $"{cell.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{styleId.Id}{EndQuoteBeginValue}" +
            $"{cell.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{FormulaCellHelper.EndQuoteBeginFormula}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.Value.StringOrPrimitive.PrimitiveValue.IntValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.TryWrite($"{cell.Value.StringOrPrimitive.PrimitiveValue.IntValue}");
    }
}
