namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetColumn
{
    bool Collapsed { get; }
    bool Hidden { get; }
    int OutlineLevel { get; }
    double Width { get; }
    IStyle Style { get; }
    IEnumerable<IWorksheetCell> Cells { get; }
}
