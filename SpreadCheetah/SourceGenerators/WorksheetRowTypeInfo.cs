namespace SpreadCheetah.SourceGenerators;

public abstract class WorksheetRowTypeInfo<T>
{
    public Func<Spreadsheet, T, CancellationToken, ValueTask> RowHandler { get; }

    private protected WorksheetRowTypeInfo(Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler)
    {
        RowHandler = rowHandler;
    }
}
