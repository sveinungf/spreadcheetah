namespace SpreadCheetah.TestHelpers.Interfaces;

public interface ISpreadsheetDocumentProperties
{
    string? Application { get; }
    string? Author { get; }
    string? Subject { get; }
    string? Title { get; }
    DateTime? Created { get; }
}
