using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration.Internal;

internal sealed class WorksheetRowTypeInfoInternal<T>(
    Func<Spreadsheet, StyleId?, CancellationToken, ValueTask> headerHandler,
    Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler,
    Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> rowRangeHandler)
    : WorksheetRowTypeInfo<T>(headerHandler, rowHandler, rowRangeHandler);