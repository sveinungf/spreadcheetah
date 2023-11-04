using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using System.Buffers;
using System.Diagnostics;
using System.IO.Compression;

#if !NET6_0_OR_GREATER
using ArgumentNullException = SpreadCheetah.Helpers.Backporting.ArgumentNullExceptionBackport;
#endif

namespace SpreadCheetah;

/// <summary>
/// The main class for generating spreadsheets with SpreadCheetah. Use <see cref="CreateNewAsync"/> to initialize a new instance.
/// </summary>
public sealed class Spreadsheet : IDisposable, IAsyncDisposable
{
    private readonly Guid _spreadsheetGuid = Guid.NewGuid();
    private readonly List<WorksheetMetadata> _worksheets = new();
    private readonly ZipArchive _archive;
    private readonly CompressionLevel _compressionLevel;
    private readonly SpreadsheetBuffer _buffer;
    private readonly bool _writeCellReferenceAttributes;
    private readonly byte[] _arrayPoolBuffer;
    private readonly NumberFormat? _defaultDateTimeFormat;
    private DefaultStyling? _defaultStyling;
    private Dictionary<ImmutableStyle, int>? _styles;
    private FileCounter? _fileCounter;
    private Worksheet? _worksheet;
    private bool _disposed;
    private bool _finished;

    private Worksheet Worksheet
    {
        get
        {
            if (_worksheet is null) ThrowHelper.NoActiveWorksheet();
            return _worksheet;
        }
    }

    private Spreadsheet(ZipArchive archive, CompressionLevel compressionLevel, int bufferSize, NumberFormat? defaultDateTimeFormat, bool writeCellReferenceAttributes)
    {
        _archive = archive;
        _compressionLevel = compressionLevel;
        _arrayPoolBuffer = ArrayPool<byte>.Shared.Rent(bufferSize);
        _buffer = new SpreadsheetBuffer(_arrayPoolBuffer);
        _defaultDateTimeFormat = defaultDateTimeFormat;
        _writeCellReferenceAttributes = writeCellReferenceAttributes;
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
        var defaultDateTimeFormat = options is null ? SpreadCheetahOptions.InitialDefaultDateTimeFormat : options.DefaultDateTimeFormat;
        var writeCellReferenceAttributes = options?.WriteCellReferenceAttributes ?? false;

        var spreadsheet = new Spreadsheet(archive, compressionLevel, bufferSize, defaultDateTimeFormat, writeCellReferenceAttributes);
        await spreadsheet.InitializeAsync(cancellationToken).ConfigureAwait(false);

        // If no style is ever added to the spreadsheet, then we can skip creating the styles.xml file.
        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        if (defaultDateTimeFormat is not null)
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
    /// The next row number for the active worksheet. The first row in a worksheet has row number 1.
    /// </summary>
    public int NextRowNumber => Worksheet.NextRowNumber;

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

        if (_worksheets.Exists(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase)))
            ThrowHelper.WorksheetNameAlreadyExists(nameof(name));

        if (_finished)
            ThrowHelper.StartWorksheetNotAllowedAfterFinish();

        return StartWorksheetInternalAsync(name, options, token);
    }

    private async ValueTask StartWorksheetInternalAsync(string name, WorksheetOptions? options, CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var path = StringHelper.Invariant($"xl/worksheets/sheet{_worksheets.Count + 1}.xml");
        var entry = _archive.CreateEntry(path, _compressionLevel);
        var entryStream = entry.Open();
        _worksheet = new Worksheet(entryStream, _defaultStyling, _buffer, _writeCellReferenceAttributes);
        await _worksheet.WriteHeadAsync(options, token).ConfigureAwait(false);
        _worksheets.Add(new WorksheetMetadata(name, path, options?.Visibility ?? WorksheetVisibility.Visible));
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
        return Worksheet.TryAddRow(cells.Span)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells)
            ? default
            : Worksheet.AddRowAsync(cells, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        if (options is null) return AddRowAsync(cells, token);
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
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
        ImmutableStyle? dateTimeStyle = _defaultDateTimeFormat is not null && style.Format is null
            ? style with { Format = _defaultDateTimeFormat }
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
    /// Note that there can be max 65534 data validations in a worksheet. This method throws a <see cref="SpreadCheetahException"/> if attempting to add more than that.
    /// </summary>
    public void AddDataValidation(string reference, DataValidation validation)
    {
        if (!TryAddDataValidation(reference, validation))
            ThrowHelper.MaxNumberOfDataValidations();
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
    /// Note that there can be max 65534 data validations in a worksheet. This method returns <c>false</c> if attempting to add more than that.
    /// </summary>
    public bool TryAddDataValidation(string reference, DataValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        ArgumentNullException.ThrowIfNull(reference);
        return Worksheet.TryAddDataValidation(reference, validation);
    }

    /// <summary>
    /// Adds a plain text note for a cell. The cell reference must be in the A1 reference style, e.g. "A1" or "C4".
    /// </summary>
    public void AddNote(string cellReference, string noteText)
    {
        ArgumentNullException.ThrowIfNull(cellReference);
        ArgumentNullException.ThrowIfNull(noteText);
        if (noteText.Length > SpreadsheetConstants.MaxNoteTextLength)
            ThrowHelper.NoteTextTooLong(nameof(noteText));

        Worksheet.AddNote(cellReference, noteText, out var firstNote);

        if (firstNote)
        {
            _fileCounter ??= new FileCounter();
            _fileCounter.WorksheetsWithNotes++;
        }
    }

    /// <summary>
    /// Merge a range of cells together. Note that only the content of the upper-left cell in the range will appear in the merged cell.
    /// The cell range must be in the A1 reference style. Some examples:
    /// <list type="bullet">
    ///   <item><term><c>A1:E5</c></term><description>References the range from cell A1 to E5.</description></item>
    ///   <item><term><c>A1:A1048576</c></term><description>References all cells in column A.</description></item>
    ///   <item><term><c>A5:XFD5</c></term><description>References all cells in row 5.</description></item>
    /// </list>
    /// </summary>
    /// <param name="cellRange"></param>
    public void MergeCells(string cellRange)
    {
        ArgumentNullException.ThrowIfNull(cellRange);

        var cellReference = CellRangeRelativeReference.Create(cellRange);
        Worksheet.MergeCells(cellReference);
    }

    public async ValueTask<EmbeddedImage> EmbedImageAsync(Stream stream, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (_finished)
            ThrowHelper.EmbedImageNotAllowedAfterFinish();
        if (!stream.CanRead)
            ThrowHelper.StreamDoesNotSupportReading(nameof(stream));
        if (_worksheet is not null)
            ThrowHelper.EmbedImageBeforeStartingWorksheet();

        const int bytesToRead = 24; // Enough to cover PNG file signature with dimensions
        using var pooledArray = await stream.ReadToPooledArrayAsync(bytesToRead, token).ConfigureAwait(false);
        var buffer = pooledArray.Memory;

        if (buffer.Length == 0)
            ThrowHelper.StreamReadNoBytes(nameof(stream));
        if (buffer.Length < bytesToRead)
            ThrowHelper.StreamReadNotEnoughBytes(nameof(stream));

        var imageType = FileSignature.GetImageTypeFromHeader(buffer.Span);
        if (imageType is null)
            ThrowHelper.StreamContentNotSupportedImageType(nameof(stream));

        _fileCounter ??= new FileCounter();
        _fileCounter.AddEmbeddedImage(imageType.Value);
        var embeddedImageId = _fileCounter.TotalEmbeddedImages;

        // TODO: What about imageCount if this fails?
        return await _archive.CreateImageEntryAsync(stream, _compressionLevel, buffer, imageType.Value, embeddedImageId, _spreadsheetGuid, token).ConfigureAwait(false);
    }

    public void AddImage(string cellReference, EmbeddedImage image, ImageOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(cellReference);
        ArgumentNullException.ThrowIfNull(image);

        if (_spreadsheetGuid != image.SpreadsheetGuid)
            ThrowHelper.CantAddImageEmbeddedInOtherSpreadsheet();

        var reference = SingleCellRelativeReference.Create(cellReference);
        var anchor = options?.GetAnchor() ?? ImageAnchor.OneCell;
        if (anchor == ImageAnchor.None)
            ThrowHelper.ResizeWithCellsNotSupportedWhenMoveWithCells(nameof(options));

        var immutableImage = new ImmutableImage(image.Id, image.Width, image.Height, anchor, options?.Size);

        _fileCounter ??= new FileCounter();
        _fileCounter.TotalAddedImages++;

        var worksheetImage = new WorksheetImage(reference, immutableImage, _fileCounter.TotalAddedImages);
        Worksheet.AddImage(worksheetImage, out var firstImage);

        if (firstImage)
            _fileCounter.WorksheetsWithImages++;
    }

    private async ValueTask FinishAndDisposeWorksheetAsync(CancellationToken token)
    {
        if (_worksheet is not { } worksheet) return;

        await worksheet.FinishAsync(token).ConfigureAwait(false);
        await worksheet.DisposeAsync().ConfigureAwait(false);

        if (worksheet.Notes is { } notes)
        {
            var sheetsWithNotes = _fileCounter?.WorksheetsWithNotes ?? 0;
            Debug.Assert(sheetsWithNotes > 0);
            using var notesPooledArray = notes.ToPooledArray();

            await CommentsXml.WriteAsync(_archive, _compressionLevel, _buffer, sheetsWithNotes, notesPooledArray.Memory, token).ConfigureAwait(false);
            await VmlDrawingXml.WriteAsync(_archive, _compressionLevel, _buffer, sheetsWithNotes, notesPooledArray.Memory, token).ConfigureAwait(false);
        }

        if (worksheet.Images is { } images)
        {
            var sheetsWithImages = _fileCounter?.WorksheetsWithImages ?? 0;
            Debug.Assert(sheetsWithImages > 0);
            await DrawingXml.WriteAsync(_archive, _compressionLevel, _buffer, sheetsWithImages, images, token).ConfigureAwait(false);
            await DrawingRelsXml.WriteAsync(_archive, _compressionLevel, _buffer, sheetsWithImages, images, token).ConfigureAwait(false);
        }

        if (_fileCounter is { } counter && (counter.WorksheetsWithImages > 0 || counter.WorksheetsWithNotes > 0))
        {
            var worksheetIndex = _worksheets.Count;
            await WorksheetRelsXml.WriteAsync(_archive, _compressionLevel, _buffer, worksheetIndex, counter, token).ConfigureAwait(false);
        }

        _worksheet = null;
    }

    /// <summary>
    /// Finalize the spreadsheet. This will write remaining metadata to the output which is important to get a valid XLSX file.
    /// No more data can be added after this has been called. Will throw a <see cref="SpreadCheetahException"/> if the spreadsheet contains no worksheets.
    /// </summary>
    public ValueTask FinishAsync(CancellationToken token = default)
    {
        if (_worksheets.Count == 0) ThrowHelper.SpreadsheetMustContainWorksheet();
        return FinishInternalAsync(token);
    }

    private async ValueTask FinishInternalAsync(CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var hasStyles = _styles != null;

        await ContentTypesXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheets, _fileCounter, hasStyles, token).ConfigureAwait(false);
        await WorkbookRelsXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheets, hasStyles, token).ConfigureAwait(false);
        await WorkbookXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheets, token).ConfigureAwait(false);

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
