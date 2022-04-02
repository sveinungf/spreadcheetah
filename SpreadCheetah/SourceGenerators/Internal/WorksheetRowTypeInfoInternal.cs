namespace SpreadCheetah.SourceGenerators.Internal;

internal sealed class WorksheetRowTypeInfoInternal<T> : WorksheetRowTypeInfo<T>
{
    public WorksheetRowTypeInfoInternal(Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler)
        : base(rowHandler)
    {
    }
}
