using SpreadCheetah.SourceGeneration.Internal;

namespace SpreadCheetah.SourceGeneration;

public static class WorksheetRowMetadataServices
{
    public static WorksheetRowTypeInfo<T> CreateObjectInfo<T>(Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler)
        => new WorksheetRowTypeInfoInternal<T>(rowHandler);
}
