namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetRow
{
    double Height { get; }
    int OutlineLevel { get; }
    IStyle Style { get; }
    IEnumerable<IWorksheetCell> Cells { get; }
}
