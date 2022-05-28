namespace SpreadCheetah.SourceGeneration.Internal;

internal sealed class WorksheetRowTypeInfoInternal<T> : WorksheetRowTypeInfo<T>
{
    public WorksheetRowTypeInfoInternal(
        Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler,
        Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> rowRangeHandler)
        : base(rowHandler, rowRangeHandler)
    {
    }
}
