namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Provides source generation-related metadata about a type.
/// The SpreadCheetah source generator will create derived types from this base class.
/// </summary>
public abstract class WorksheetRowTypeInfo<T>
{
    /// <summary>
    /// Method for adding a row to a worksheet from a type.
    /// </summary>
    public Func<Spreadsheet, T, CancellationToken, ValueTask> RowHandler { get; }

    private protected WorksheetRowTypeInfo(Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler)
    {
        RowHandler = rowHandler;
    }
}
