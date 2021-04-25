using SpreadCheetah.CellValueWriters.Boolean;
using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellValueWriters.String;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters
{
    internal abstract class CellValueWriter
    {
        public static CellValueWriter NullNumber { get; } = new NullNumberCellValueWriter();
        public static CellValueWriter NullBoolean { get; } = new NullBooleanCellValueWriter();
        public static CellValueWriter NullString { get; } = new NullStringCellValueWriter();
        public static CellValueWriter Integer { get; } = new IntegerCellValueWriter();
        public static CellValueWriter Float { get; } = new FloatCellValueWriter();
        public static CellValueWriter Double { get; } = new DoubleCellValueWriter();
        public static CellValueWriter TrueBoolean { get; } = new TrueBooleanCellValueWriter();
        public static CellValueWriter FalseBoolean { get; } = new FalseBooleanCellValueWriter();
        public static CellValueWriter String { get; } = new StringCellValueWriter();

        public abstract bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer);
        public abstract bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer);
        public abstract bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer);
        public abstract bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer, out int bytesNeeded);
        public abstract bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer, out int bytesNeeded);
        public abstract bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer, out int bytesNeeded);
        public abstract bool WriteStartElement(in DataCell cell, SpreadsheetBuffer buffer);
        public abstract bool WriteStartElement(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer);
        public abstract bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer);
        public abstract bool TryWriteEndElement(in DataCell cell, SpreadsheetBuffer buffer);
        public abstract bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer);
    }
}
