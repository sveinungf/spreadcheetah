using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Provides source generation-related metadata about a type.
/// The SpreadCheetah source generator will create derived types from this base class.
/// </summary>
public abstract class WorksheetRowTypeInfo<T>
{
    private readonly Func<WorksheetOptions>? _worksheetOptionsFactory;
    private WorksheetOptions? _worksheetOptionsInstance;

    /// <summary>
    /// A cached instance of <see cref="WorksheetOptions"/>. The instance should not be mutated.
    /// It is only accessible internally because the type is mutable.
    /// </summary>
    internal WorksheetOptions WorksheetOptionsInstance => _worksheetOptionsInstance ??= CreateWorksheetOptions();

    /// <summary>
    /// Creates a new instance of <see cref="WorksheetOptions"/> with column widths set by <see cref="ColumnWidthAttribute"/>.
    /// </summary>
    public WorksheetOptions CreateWorksheetOptions() => _worksheetOptionsFactory?.Invoke() ?? new();

    public Func<Spreadsheet, WorksheetRowDependencyInfo>? CreateWorksheetRowDependencyInfo { get; }

    /// <summary>
    /// Method for adding a header to a worksheet from a type.
    /// </summary>
    public Func<Spreadsheet, StyleId?, CancellationToken, ValueTask> HeaderHandler { get; }

    /// <summary>
    /// Method for adding a row to a worksheet from a type.
    /// </summary>
    public Func<Spreadsheet, T, CancellationToken, ValueTask> RowHandler { get; }

    /// <summary>
    /// Method for adding a range of rows to a worksheet from a type.
    /// </summary>
    public Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> RowRangeHandler { get; }

    private protected WorksheetRowTypeInfo(
        Func<Spreadsheet, StyleId?, CancellationToken, ValueTask> headerHandler,
        Func<Spreadsheet, T, CancellationToken, ValueTask> rowHandler,
        Func<Spreadsheet, IEnumerable<T>, CancellationToken, ValueTask> rowRangeHandler,
        Func<WorksheetOptions>? worksheetOptionsFactory,
        Func<Spreadsheet, WorksheetRowDependencyInfo>? createWorksheetRowDependencyInfo)
    {
        _worksheetOptionsFactory = worksheetOptionsFactory;
        CreateWorksheetRowDependencyInfo = createWorksheetRowDependencyInfo;
        HeaderHandler = headerHandler;
        RowHandler = rowHandler;
        RowRangeHandler = rowRangeHandler;
    }
}
