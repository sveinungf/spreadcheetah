using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters
{
    internal abstract class CellValueWriter
    {
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
