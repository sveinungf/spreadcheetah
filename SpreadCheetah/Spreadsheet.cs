using SpreadCheetah.DataValidation;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System.Buffers;
using System.IO.Compression;

namespace SpreadCheetah;

/// <summary>
/// The main class for generating spreadsheets with SpreadCheetah. Use <see cref="CreateNewAsync"/> to initialize a new instance.
/// </summary>
public sealed class Spreadsheet : IDisposable, IAsyncDisposable
{
    // Invalid worksheet name characters in Excel
    private const string InvalidSheetNameCharString = "/\\*?[]";
    private static readonly char[] InvalidSheetNameChars = InvalidSheetNameCharString.ToCharArray();

    private readonly List<string> _worksheetNames = new();
    private readonly List<string> _worksheetPaths = new();
    private readonly ZipArchive _archive;
    private readonly CompressionLevel _compressionLevel;
    private readonly SpreadsheetBuffer _buffer;
    private readonly byte[] _arrayPoolBuffer;
    private List<Style>? _styles;
    private Worksheet? _worksheet;
    private bool _disposed;
    private bool _finished;

    private Worksheet Worksheet => _worksheet ?? throw new SpreadCheetahException("There is no active worksheet.");

    private Spreadsheet(ZipArchive archive, CompressionLevel compressionLevel, int bufferSize)
    {
        _archive = archive;
        _compressionLevel = compressionLevel;
        _arrayPoolBuffer = ArrayPool<byte>.Shared.Rent(bufferSize);
        _buffer = new SpreadsheetBuffer(_arrayPoolBuffer);
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
        options ??= new SpreadCheetahOptions();
        var spreadsheet = new Spreadsheet(archive, GetCompressionLevel(options.CompressionLevel), options.BufferSize);
        await spreadsheet.InitializeAsync(cancellationToken).ConfigureAwait(false);
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
        if (name is null)
            throw new ArgumentNullException(nameof(name));

        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("The name can not be empty or consist only of whitespace.", nameof(name));

        if (name.Length > 31)
            throw new ArgumentException("The name can not be more than 31 characters.", nameof(name));

        if (name.StartsWith("\'", StringComparison.Ordinal) || name.EndsWith("\'", StringComparison.Ordinal))
            throw new ArgumentException("The name can not start or end with a single quote.", nameof(name));

        if (name.IndexOfAny(InvalidSheetNameChars) != -1)
            throw new ArgumentException("The name can not contain any of the following characters: " + InvalidSheetNameCharString, nameof(name));

        if (_worksheetNames.Contains(name, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException("A worksheet with the given name already exists.", nameof(name));

        if (_finished)
            throw new SpreadCheetahException("Can't start another worksheet after " + nameof(FinishAsync) + " has been called.");

        return StartWorksheetInternalAsync(name, options, token);
    }

    private async ValueTask StartWorksheetInternalAsync(string name, WorksheetOptions? options, CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var path = $"xl/worksheets/sheet{_worksheetNames.Count + 1}.xml";
        var entry = _archive.CreateEntry(path, _compressionLevel);
        var entryStream = entry.Open();
        _worksheet = new Worksheet(entryStream, _buffer);
        await _worksheet.WriteHeadAsync(options, token).ConfigureAwait(false);
        _worksheetNames.Add(name);
        _worksheetPaths.Add(path);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(Cell[] cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(DataCell[] cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(StyledCell[] cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells.Slice(currentIndex), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells.Slice(currentIndex), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells.Slice(currentIndex), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
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
        ThrowIfNull(cells, nameof(cells));
        return Worksheet.TryAddRow(cells, options, out var rowStartWritten, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, rowStartWritten, currentIndex, cells.Count, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
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
        ThrowIfNull(cells, nameof(cells));
        return Worksheet.TryAddRow(cells, options, out var rowStartWritten, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, rowStartWritten, currentIndex, cells.Count, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, CancellationToken token = default)
    {
        ThrowIfNull(cells, nameof(cells));
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
        ThrowIfNull(cells, nameof(cells));
        return Worksheet.TryAddRow(cells, options, out var rowStartWritten, out var currentIndex)
            ? default
            : Worksheet.AddRowAsync(cells, options, rowStartWritten, currentIndex, cells.Count, token);
    }

    private static void ThrowIfNull<T>(T? obj, string paramName)
    {
        if (obj is null)
            throw new ArgumentNullException(paramName);
    }

    /// <summary>
    /// Adds a reusable style to the spreadsheet and returns a style ID.
    /// </summary>
    public StyleId AddStyle(Style style)
    {
        ThrowIfNull(style, nameof(style));

        if (_styles is null)
            _styles = new List<Style>();

        _styles.Add(style);
        return new StyleId(_styles.Count);
    }

    public void AddDataValidation(string reference, BaseValidation validation)
    {
        ThrowIfNull(validation, nameof(validation));
        Worksheet.AddDataValidation(reference, validation);
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

        if (_styles != null)
            await StylesXml.WriteAsync(_archive, _compressionLevel, _buffer, _styles, token).ConfigureAwait(false);

        _finished = true;
    }

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_worksheet != null)
            await _worksheet.DisposeAsync().ConfigureAwait(false);

        ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
        _archive?.Dispose();
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _worksheet?.Dispose();
        ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
        _archive?.Dispose();
    }
}
