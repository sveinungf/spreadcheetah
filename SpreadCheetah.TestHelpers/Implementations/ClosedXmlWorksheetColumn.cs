using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetColumn(IXLColumn column) : IWorksheetColumn
{
    public bool Collapsed
    {
        get
        {
            var property = column.GetType().GetProperty("Collapsed")
                ?? throw new InvalidOperationException($"The property 'Collapsed' does not exist on type '{column.GetType().FullName}'.");

            var value = property.GetValue(column);
            return value is not bool boolValue
                ? throw new InvalidOperationException($"The property 'Collapsed' on type '{column.GetType().FullName}' is not of type 'bool'.")
                : boolValue;
        }
    }

    public bool Hidden => column.IsHidden;

    public int OutlineLevel => column.OutlineLevel;

    public double Width => column.Width + 0.71062;

    public IStyle Style => ClosedXmlStyle.Create(column.Style);

    public IEnumerable<IWorksheetCell> Cells => column.CellsUsed().Select(x => new ClosedXmlWorksheetCell(x));
}
