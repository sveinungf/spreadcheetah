using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.SourceGeneration.Internal;

internal sealed class WorksheetRowTypeInfoInternal<T>(
    Func<Spreadsheet, StyleId?, CancellationToken, ValueTask> headerHandler,
    Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler,
    Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> rowRangeHandler,
    Func<WorksheetOptions>? worksheetOptionsFactory,
    Func<Spreadsheet, WorksheetRowDependencyInfo>? createWorksheetRowDependencyInfo)
    : WorksheetRowTypeInfo<T>(headerHandler, rowHandler, rowRangeHandler, worksheetOptionsFactory, createWorksheetRowDependencyInfo);