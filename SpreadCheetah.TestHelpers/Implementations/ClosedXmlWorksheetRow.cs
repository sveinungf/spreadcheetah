using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetRow(IXLRow row) : IWorksheetRow
{
    public bool Collapsed
    {
        get
        {
            var property = row.GetType().GetProperty("Collapsed")
                ?? throw new InvalidOperationException($"The property 'Collapsed' does not exist on type '{row.GetType().FullName}'.");

            var value = property.GetValue(row);
            return value is not bool boolValue
                ? throw new InvalidOperationException($"The property 'Collapsed' on type '{row.GetType().FullName}' is not of type 'bool'.")
                : boolValue;
        }
    }

    public bool Hidden => row.IsHidden;

    public double Height => row.Height;

    public int OutlineLevel => row.OutlineLevel;

    public IStyle Style => ClosedXmlStyle.Create(row.Style);

    public IEnumerable<IWorksheetCell> Cells => row.IsEmpty()
        ? []
        : row.Cells(false).Select(x => new ClosedXmlWorksheetCell(x));
}
