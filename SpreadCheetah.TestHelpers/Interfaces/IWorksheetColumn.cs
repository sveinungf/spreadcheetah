namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetColumn
{
    bool Hidden { get; }
    double Width { get; }
    IStyle Style { get; }
    IEnumerable<IWorksheetCell> Cells { get; }
}
