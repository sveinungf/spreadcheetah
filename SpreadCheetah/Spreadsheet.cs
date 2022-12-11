using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using System.Buffers;
using System.IO.Compression;

namespace SpreadCheetah;

/// <summary>
/// The main class for generating spreadsheets with SpreadCheetah. Use <see cref="CreateNewAsync"/> to initialize a new instance.
/// </summary>
public sealed class Spreadsheet : IDisposable, IAsyncDisposable
{
    private readonly List<string> _worksheetNames = new();
    private readonly List<string> _worksheetPaths = new();
    private readonly ZipArchive _archive;
    private readonly CompressionLevel _compressionLevel;
    private readonly SpreadsheetBuffer _buffer;
    private readonly byte[] _arrayPoolBuffer;
    private readonly string? _defaultDateTimeNumberFormat;
    private DefaultStyling? _defaultStyling;
    private Dictionary<ImmutableStyle, int>? _styles;
    private Worksheet? _worksheet;
    private bool _disposed;
    private bool _finished;

    private Worksheet Worksheet => _worksheet ?? throw new SpreadCheetahException("There is no active worksheet.");

    private Spreadsheet(ZipArchive archive, CompressionLevel compressionLevel, int bufferSize, string? defaultDateTimeNumberFormat)
    {
        _archive = archive;
        _compressionLevel = compressionLevel;
        _arrayPoolBuffer = ArrayPool<byte>.Shared.Rent(bufferSize);
        _buffer = new SpreadsheetBuffer(_arrayPoolBuffer);
        _defaultDateTimeNumberFormat = defaultDateTimeNumberFormat;
    }

    /// <summary>
    /// Initializes a new <see cref="Spreadsheet"/> that writes its output to a <see cref="Stream"/>.
    /// </summary>
    public static async ValueTask<Spreadsheet> CreateNewAsync(
        Stream stream,
        SpreadCheetahOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        var archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
        var bufferSize = options?.BufferSize ?? SpreadCheetahOptions.DefaultBufferSize;
        var compressionLevel = GetCompressionLevel(options?.CompressionLevel ?? SpreadCheetahOptions.DefaultCompressionLevel);
        var defaultDateTimeNumberFormat = options is null ? NumberFormats.DateTimeSortable : options.DefaultDateTimeNumberFormat;

        var spreadsheet = new Spreadsheet(archive, compressionLevel, bufferSize, defaultDateTimeNumberFormat);
        await spreadsheet.InitializeAsync(cancellationToken).ConfigureAwait(false);

        // If no style is ever added to the spreadsheet, then we can skip creating the styles.xml file.
        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        if (defaultDateTimeNumberFormat is not null)
            spreadsheet.AddDefaultStyle();

        return spreadsheet;
    }

    private static CompressionLevel GetCompressionLevel(SpreadCheetahCompressionLevel level)
    {
        return level == SpreadCheetahCompressionLevel.Optimal ? CompressionLevel.Optimal : CompressionLevel.Fastest;
    }

    private ValueTask InitializeAsync(CancellationToken token)
    {
        return RelationshipsXml.WriteAsync(_archive, _compressionLevel, _buffer, token);
    }

    /// <summary>
    /// Starts a new worksheet in the spreadsheet. Every spreadsheet must have at least one worksheet.
    /// The name must satisfy these requirements:
    /// <list type="bullet">
    ///   <item><description>Can not be empty or consist only of whitespace.</description></item>
    ///   <item><description>Can not be more than 31 characters.</description></item>
    ///   <item><description>Can not start or end with a single quote.</description></item>
    ///   <item><description>Can not contain these characters: / \ * ? [ ] </description></item>
    ///   <item><description>Must be unique across all worksheets.</description></item>
    /// </list>
    /// </summary>
    public ValueTask StartWorksheetAsync(string name, WorksheetOptions? options = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(name);
        name.EnsureValidWorksheetName();

        if (_worksheetNames.Contains(name, StringComparer.OrdinalIgnoreCase))
            ThrowHelper.WorksheetNameAlreadyExists(nameof(name));

        if (_finished)
            ThrowHelper.StartWorksheetNotAllowedAfterFinish();

        return StartWorksheetInternalAsync(name, options, token);
    }

    private async ValueTask StartWorksheetInternalAsync(string name, WorksheetOptions? options, CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var path = $"xl/worksheets/sheet{_worksheetNames.Count + 1}.xml";
        var entry = _archive.CreateEntry(path, _compressionLevel);
        var entryStream = entry.Open();
        _worksheet = new Worksheet(entryStream, _defaultStyling, _buffer);
        await _worksheet.WriteHeadAsync(options, token).ConfigureAwait(false);
        _worksheetNames.Add(name);
        _worksheetPaths.Add(path);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(Cell[] cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(Cell[] cells, RowOptions? options, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(DataCell[] cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(DataCell[] cells, RowOptions? options, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(StyledCell[] cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(StyledCell[] cells, RowOptions? options, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, currentIndex, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, currentIndex, token);
    }

    /// <summary>
    /// Add object as a row in the active worksheet.
    /// Each property with a public getter on the object will be added as a cell in the row.
    /// The <see cref="WorksheetRowTypeInfo{T}"/> type must be generated by a source generator.
    /// </summary>
    public ValueTask AddAsRowAsync<T>(T obj, WorksheetRowTypeInfo<T> typeInfo, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(typeInfo);
        return typeInfo.RowHandler(this, obj, token);
    }

    /// <summary>
    /// Add objects as rows in the active worksheet.
    /// Each property with a public getter on the object will be added as a cell in the row.
    /// The <see cref="WorksheetRowTypeInfo{T}"/> type must be generated by a source generator.
    /// </summary>
    public ValueTask AddRangeAsRowsAsync<T>(IEnumerable<T> objs, WorksheetRowTypeInfo<T> typeInfo, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(objs);
        ArgumentNullException.ThrowIfNull(typeInfo);
        return typeInfo.RowRangeHandler(this, objs, token);
    }

    /// <summary>
    /// Adds a reusable style to the spreadsheet and returns a style ID.
    /// </summary>
    public StyleId AddStyle(Style style)
    {
        ArgumentNullException.ThrowIfNull(style);

        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        if (_styles is null)
            AddDefaultStyle();

        return AddStyle(ImmutableStyle.From(style));
    }

    private void AddDefaultStyle()
    {
        var defaultFont = new ImmutableFont(null, false, false, false, Font.DefaultSize, null);
        var defaultStyle = new ImmutableStyle(new ImmutableAlignment(), new ImmutableBorder(), new ImmutableFill(), defaultFont, null);
        var styleId = AddStyle(defaultStyle);

        if (styleId.Id != styleId.DateTimeId)
            _defaultStyling = new DefaultStyling(styleId.DateTimeId);
    }

    private StyleId AddStyle(in ImmutableStyle style)
    {
        // Optionally add another style for DateTime when there is no explicit number format in the new style.
        ImmutableStyle? dateTimeStyle = _defaultDateTimeNumberFormat is not null && style.NumberFormat is null
            ? style with { NumberFormat = _defaultDateTimeNumberFormat }
            : null;

        _styles ??= new Dictionary<ImmutableStyle, int>();
        if (_styles.TryGetValue(style, out var id))
        {
            return dateTimeStyle is not null && _styles.TryGetValue(dateTimeStyle.Value, out var dateTimeId)
                ? new StyleId(id, dateTimeId)
                : new StyleId(id, id);
        }

        var newId = _styles.Count;
        _styles[style] = newId;

        if (dateTimeStyle is null)
            return new StyleId(newId, newId);

        var newDateTimeId = newId + 1;
        _styles[dateTimeStyle.Value] = newDateTimeId;
        return new StyleId(newId, newDateTimeId);
    }

    /// <summary>
    /// Adds data validation for a cell or a range of cells. The reference must be in the A1 reference style. Some examples:
    /// <list type="bullet">
    ///   <item><term><c>A1</c></term><description>References the top left cell.</description></item>
    ///   <item><term><c>C4</c></term><description>References the cell in column C row 4.</description></item>
    ///   <item><term><c>A1:E5</c></term><description>References the range from cell A1 to E5.</description></item>
    ///   <item><term><c>A1:A1048576</c></term><description>References all cells in column A.</description></item>
    ///   <item><term><c>A5:XFD5</c></term><description>References all cells in row 5.</description></item>
    /// </list>
    /// </summary>
    public void AddDataValidation(string reference, DataValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        ArgumentNullException.ThrowIfNull(reference);

        var cellReference = CellReference.Create(reference);
        Worksheet.AddDataValidation(cellReference, validation);
    }

    private async ValueTask FinishAndDisposeWorksheetAsync(CancellationToken token)
    {
        if (_worksheet is null) return;
        await _worksheet.FinishAsync(token).ConfigureAwait(false);
        await _worksheet.DisposeAsync().ConfigureAwait(false);
        _worksheet = null;
    }

    /// <summary>
    /// Finalize the spreadsheet. This will write remaining metadata to the output which is important to get a valid XLSX file.
    /// No more data can be added after this has been called. Will throw a <see cref="SpreadCheetahException"/> if the spreadsheet contains no worksheets.
    /// </summary>
    public ValueTask FinishAsync(CancellationToken token = default)
    {
        if (_worksheetNames.Count == 0)
            throw new SpreadCheetahException("Spreadsheet must contain at least one worksheet.");

        return FinishInternalAsync(token);
    }

    private async ValueTask FinishInternalAsync(CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var hasStyles = _styles != null;
        await ContentTypesXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetPaths, hasStyles, token).ConfigureAwait(false);
        await WorkbookRelsXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetPaths, hasStyles, token).ConfigureAwait(false);
        await WorkbookXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetNames, token).ConfigureAwait(false);

        if (_styles is not null)
            await StylesXml.WriteAsync(_archive, _compressionLevel, _buffer, _styles, token).ConfigureAwait(false);

        _finished = true;

        // The XLSX can become corrupt if the archive is not flushed before the resulting stream is being used.
        // ZipArchive.Dispose() is currently (up to .NET 7) the only way to flush the archive to the resulting stream.
        // In case the user would use the resulting stream before Spreadsheet.Dispose(), the flush must happen here to prevent a corrupt XLSX.
        _archive.Dispose();
    }

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_worksheet != null)
            await _worksheet.DisposeAsync().ConfigureAwait(false);

        ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _worksheet?.Dispose();
        ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
    }
}
