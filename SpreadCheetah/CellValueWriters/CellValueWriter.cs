using SpreadCheetah.CellValueWriters.Boolean;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellValueWriters.Time;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.CellValueWriters;

internal abstract class CellValueWriter
{
    private static readonly CellValueWriter[] Writers =
    [
        new NullValueWriter(),
        new StringCellValueWriter(),
        new IntegerCellValueWriter(),
        new FloatCellValueWriter(),
        new DoubleCellValueWriter(),
        new NullDateTimeCellValueWriter(),
        new DateTimeCellValueWriter(),
        new FalseBooleanCellValueWriter(),
        new TrueBooleanCellValueWriter(),
        new ReadOnlyMemoryOfCharCellValueWriter()
    ];

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static CellValueWriter GetWriter(CellWriterType type) => Writers[(int)type];

    public abstract bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer);
    public abstract bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer);
    public abstract bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer);
    public abstract bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state);
    public abstract bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state);
    public abstract bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state);
    public abstract bool WriteStartElement(StyleId? styleId, SpreadsheetBuffer buffer);
    public abstract bool WriteStartElementWithReference(StyleId? styleId, CellWriterState state);
    public abstract bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer);
    public abstract bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state);
    public abstract bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex);
    public abstract bool TryWriteEndElement(SpreadsheetBuffer buffer);
    public abstract bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer);
}
