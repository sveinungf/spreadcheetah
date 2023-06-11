using SpreadCheetah.CellValueWriters.Boolean;
using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellValueWriters.Time;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters;

internal abstract class CellValueWriter
{
    public static CellValueWriter Null { get; } = new NullValueWriter();
    public static CellValueWriter Integer { get; } = new IntegerCellValueWriter();
    public static CellValueWriter Float { get; } = new FloatCellValueWriter();
    public static CellValueWriter Double { get; } = new DoubleCellValueWriter();
    public static CellValueWriter DateTime { get; } = new DateTimeCellValueWriter();
    public static CellValueWriter NullDateTime { get; } = new NullDateTimeCellValueWriter();
    public static CellValueWriter TrueBoolean { get; } = new TrueBooleanCellValueWriter();
    public static CellValueWriter FalseBoolean { get; } = new FalseBooleanCellValueWriter();
    public static CellValueWriter String { get; } = new StringCellValueWriter();

    public abstract bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state);
    public abstract bool TryWriteCell(in DataCell cell, StyleId styleId, CellWriterState state);
    public abstract bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state);
    public abstract bool WriteStartElement(SpreadsheetBuffer buffer);
    public abstract bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer);
    public abstract bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer);
    public abstract bool CanWriteValuePieceByPiece(in DataCell cell);
    public abstract bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex);
    public abstract bool TryWriteEndElement(SpreadsheetBuffer buffer);
    public abstract bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer);
    public abstract bool Equals(in CellValue value, in CellValue other);
    public abstract int GetHashCodeFor(in CellValue value);

    protected static bool TryWriteCellStartWithReference(CellWriterState state, Span<byte> bytes, ref int bytesWritten)
    {
        var written = 0;

        if (!"<c r=\""u8.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWriteCellReference(state.Column + 1, state.NextRowIndex - 1, bytes, ref written)) return false;

        bytesWritten += written;
        return true;
    }
}
