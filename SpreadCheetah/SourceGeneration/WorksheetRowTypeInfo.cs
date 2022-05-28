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

    /// <summary>
    /// Method for adding a range of rows to a worksheet from a type.
    /// </summary>
    public Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> RowRangeHandler { get; }

    private protected WorksheetRowTypeInfo(
        Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler,
        Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> rowRangeHandler)
    {
        RowHandler = rowHandler;
        RowRangeHandler = rowRangeHandler;
    }
}
