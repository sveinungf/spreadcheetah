namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetRow
{
    bool Collapsed { get; }
    bool Hidden { get; }
    double Height { get; }
    int OutlineLevel { get; }
    IStyle Style { get; }
    IEnumerable<IWorksheetCell> Cells { get; }
}
