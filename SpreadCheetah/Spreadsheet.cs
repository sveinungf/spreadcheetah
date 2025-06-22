using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using SpreadCheetah.Metadata;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.MetadataXml.Styles;
using SpreadCheetah.MetadataXml.Tables;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Tables;
using SpreadCheetah.Tables.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;

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
    private readonly List<WorksheetMetadata> _worksheets = new(1);
    private readonly ZipArchiveManager _zipArchiveManager;
    private readonly SpreadsheetBuffer _buffer;
    private readonly DocumentProperties? _documentProperties;
    private readonly bool _writeCellReferenceAttributes;
    private Dictionary<Type, WorksheetRowDependencyInfo>? _worksheetRowDependencyInfo;
    private FileCounter? _fileCounter;
    private StyleManager? _styleManager;
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

    private Spreadsheet(ZipArchiveManager zipArchiveManager, int bufferSize, NumberFormat? defaultDateTimeFormat,
        DocumentProperties? documentProperties, bool writeCellReferenceAttributes)
    {
        _zipArchiveManager = zipArchiveManager;
        _buffer = new SpreadsheetBuffer(bufferSize);
        _documentProperties = documentProperties;
        _writeCellReferenceAttributes = writeCellReferenceAttributes;

        // If no style is ever added to the spreadsheet, then we can skip creating the styles.xml file.
        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        if (defaultDateTimeFormat is { } format)
            _styleManager = new(format);
    }

    /// <summary>
    /// Initializes a new <see cref="Spreadsheet"/> that writes its output to a <see cref="Stream"/>.
    /// </summary>
    public static ValueTask<Spreadsheet> CreateNewAsync(
        Stream stream,
        SpreadCheetahOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        var zipArchiveManager = new ZipArchiveManager(stream, options?.CompressionLevel);
        var bufferSize = options?.BufferSize ?? SpreadCheetahOptions.DefaultBufferSize;
        var defaultDateTimeFormat = options is null ? SpreadCheetahOptions.InitialDefaultDateTimeFormat : options.DefaultDateTimeFormat;
        var writeCellReferenceAttributes = options?.WriteCellReferenceAttributes ?? false;

        var documentProperties = options switch
        {
            { DocumentProperties: { } docProps } => docProps with { },
            { DocumentProperties: null } => null,
            null => DocumentProperties.Default
        };

#pragma warning disable CA2000 // Dispose objects before losing scope
        var spreadsheet = new Spreadsheet(zipArchiveManager, bufferSize, defaultDateTimeFormat, documentProperties, writeCellReferenceAttributes);
#pragma warning restore CA2000 // Dispose objects before losing scope
        _ = cancellationToken;
        return new ValueTask<Spreadsheet>(spreadsheet);
    }

    /// <summary>
    /// The next row number for the active worksheet. The first row in a worksheet has row number 1.
    /// </summary>
    public int NextRowNumber => Worksheet.NextRowNumber;

    /// <summary>
    /// <para>
    /// Starts a new worksheet in the spreadsheet. Every spreadsheet must have at least one worksheet.
    /// </para>
    /// <para>
    /// The worksheet name must satisfy these requirements:
    /// <list type="bullet">
    ///   <item><description>Can not be empty or consist only of whitespace.</description></item>
    ///   <item><description>Can not be more than 31 characters.</description></item>
    ///   <item><description>Can not start or end with a single quote.</description></item>
    ///   <item><description>Can not contain these characters: / \ * ? [ ] </description></item>
    ///   <item><description>Must be unique across all worksheets.</description></item>
    /// </list>
    /// </para>
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

    /// <summary>
    /// <para>
    /// Starts a new worksheet in the spreadsheet. Every spreadsheet must have at least one worksheet.
    /// This overload takes a <see cref="WorksheetRowTypeInfo{T}"/> type generated by the source generator,
    /// and the worksheet will be created with column widths set by <see cref="ColumnWidthAttribute"/>.
    /// </para>
    /// <para>
    /// The worksheet name must satisfy these requirements:
    /// <list type="bullet">
    ///   <item><description>Can not be empty or consist only of whitespace.</description></item>
    ///   <item><description>Can not be longer than 31 characters.</description></item>
    ///   <item><description>Can not start or end with a single quote.</description></item>
    ///   <item><description>Can not contain these characters: / \ * ? [ ] </description></item>
    ///   <item><description>Must be unique across all worksheets.</description></item>
    /// </list>
    /// </para>
    /// </summary>
    public ValueTask StartWorksheetAsync<T>(string name, WorksheetRowTypeInfo<T> typeInfo, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(typeInfo);
        return StartWorksheetAsync(name, typeInfo.WorksheetOptionsInstance, token);
    }

    private async ValueTask StartWorksheetInternalAsync(string name, WorksheetOptions? options, CancellationToken token)
    {
        await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

        var path = StringHelper.Invariant($"xl/worksheets/sheet{_worksheets.Count + 1}.xml");
        var entryStream = _zipArchiveManager.OpenEntry(path);
        _worksheet = new Worksheet(entryStream, _styleManager?.DefaultStyling, _buffer, _writeCellReferenceAttributes, options);
        _worksheets.Add(new WorksheetMetadata(name, path, options?.Visibility ?? WorksheetVisibility.Visible));

        await WorksheetStartXml.WriteAsync(options, _buffer, entryStream, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Starts a table from the next row. To add headers to the table, make a subsequent call to
    /// <see cref="AddHeaderRowAsync(IList{string}, StyleId?, CancellationToken)"/> or one of its overloads.
    /// The table will get a total row depending on the column options set on the table.
    /// The total row is written when the table is finished. The table automatically finishes when the
    /// worksheet finishes, but you can also explicitly finish a table with <see cref="FinishTableAsync(CancellationToken)"/>.
    /// Currently there can only be one active table at a time.
    /// </summary>
    public void StartTable(Table table, string firstColumnName = "A")
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentNullException.ThrowIfNull(firstColumnName);

        if (!SpreadsheetUtility.TryParseColumnName(firstColumnName.AsSpan(), out var firstColumnNumber))
            ThrowHelper.ColumnNameInvalid(nameof(firstColumnName));

        var fileCounter = _fileCounter ??= new FileCounter();

        var tableName = table.Name;
        if (tableName is null)
            tableName = fileCounter.AssignTableNameForCurrentWorksheet();
        else if (!fileCounter.TryAssignTableNameForCurrentWorksheet(tableName))
            TableThrowHelper.NameAlreadyExists(nameof(table));

        var immutableTable = ImmutableTable.From(table, tableName);
        Worksheet.StartTable(immutableTable, firstColumnNumber);
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
    [OverloadResolutionPriority(1)]
    public ValueTask AddRowAsync(DataCell[] cells, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return AddRowAsync(cells.AsMemory(), token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    [OverloadResolutionPriority(1)]
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
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, CancellationToken token = default)
    {
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, CancellationToken token = default)
    {
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        return Worksheet.TryAddRow(cells.Span, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken token = default)
    {
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions? options, CancellationToken token = default)
    {
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
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions? options, CancellationToken token = default)
    {
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
        return AddRowAsync(cells, null, token);
    }

    /// <summary>
    /// Adds a row of cells to the worksheet and increments the current row number by 1.
    /// </summary>
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions? options, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return Worksheet.TryAddRow(cells, options)
            ? default
            : Worksheet.AddRowAsync(cells, options, token);
    }

    /// <summary>
    /// Add a row of header names in the active worksheet. A style can optionally be applied to all the cells in the row.
    /// </summary>
    public ValueTask AddHeaderRowAsync(string[] headerNames, StyleId? styleId = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(headerNames);
        return AddHeaderRowAsync(headerNames.AsMemory(), styleId, token);
    }

    /// <summary>
    /// Add a row of header names in the active worksheet. A style can optionally be applied to all the cells in the row.
    /// </summary>
    public async ValueTask AddHeaderRowAsync(ReadOnlyMemory<string> headerNames, StyleId? styleId = null, CancellationToken token = default)
    {
        var headerNamesSpan = headerNames.Span;
        var activeTable = Worksheet.GetActiveTable();
        if (activeTable?.FirstRow == Worksheet.NextRowNumber)
            activeTable.SetHeaderNames(headerNamesSpan);

        using var cells = PooledArray<StyledCell>.Create(headerNamesSpan.Length);
        headerNamesSpan.CopyToCells(cells.Span, styleId);
        await AddRowAsync(cells.Memory, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Add a row of header names in the active worksheet. A style can optionally be applied to all the cells in the row.
    /// </summary>
    public async ValueTask AddHeaderRowAsync(IList<string> headerNames, StyleId? styleId = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(headerNames);
        using var pooledArray = headerNames.ToPooledArray();
        await AddHeaderRowAsync(pooledArray.Memory, styleId, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Add a row of header names in the active worksheet.
    /// This functionality depends on the source generator, which will generate the <see cref="WorksheetRowTypeInfo{T}"/> type.
    /// For properties of <typeparamref name="T"/>, the header name from the <see cref="ColumnHeaderAttribute"/> attribute will be used when set,
    /// otherwise the property name will be used instead.
    /// A style can optionally be applied to all the cells in the row.
    /// </summary>
    public ValueTask AddHeaderRowAsync<T>(WorksheetRowTypeInfo<T> typeInfo, StyleId? styleId = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(typeInfo);
        return typeInfo.HeaderHandler(this, styleId, token);
    }

    /// <summary>
    /// Add object as a row in the active worksheet.
    /// This functionality depends on the source generator, which will generate the <see cref="WorksheetRowTypeInfo{T}"/> type.
    /// Each property with a public getter on the object will be added as a cell in the row.
    /// </summary>
    public ValueTask AddAsRowAsync<T>(T obj, WorksheetRowTypeInfo<T> typeInfo, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(typeInfo);
        return typeInfo.RowHandler(this, obj, token);
    }

    /// <summary>
    /// Add objects as rows in the active worksheet.
    /// This functionality depends on the source generator, which will generate the <see cref="WorksheetRowTypeInfo{T}"/> type.
    /// Each property with a public getter on the object will be added as a cell in the row.
    /// </summary>
    public ValueTask AddRangeAsRowsAsync<T>(IEnumerable<T> objs, WorksheetRowTypeInfo<T> typeInfo, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(objs);
        ArgumentNullException.ThrowIfNull(typeInfo);
        return typeInfo.RowRangeHandler(this, objs, token);
    }

    /// <summary>
    /// Used by code generated by the source generator and is not intended to be used directly.
    /// This method is used to cache dependencies to avoid redundant lookups.
    /// The return type is immutable and the cache will return the same instance once it has been created.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public WorksheetRowDependencyInfo GetOrCreateWorksheetRowDependencyInfo<T>(WorksheetRowTypeInfo<T> typeInfo)
    {
        var dictionary = _worksheetRowDependencyInfo ??= [];
        if (dictionary.TryGetValue(typeof(T), out var result))
            return result;

        ArgumentNullException.ThrowIfNull(typeInfo);
        result = typeInfo.CreateWorksheetRowDependencyInfo?.Invoke(this) ?? new([]);
        dictionary[typeof(T)] = result;
        return result;
    }

    /// <summary>
    /// Adds a reusable style to the spreadsheet and returns a style ID.
    /// </summary>
    public StyleId AddStyle(Style style)
    {
        ArgumentNullException.ThrowIfNull(style);

        var styleManager = _styleManager ??= new(defaultDateTimeFormat: null);
        return styleManager.AddStyleIfNotExists(ImmutableStyle.From(style));
    }

    /// <summary>
    /// <para>
    /// Adds a named style to the spreadsheet and returns a style ID.
    /// This also allows for calling <see cref="GetStyleId(string)"/> with the name to get the style ID.
    /// </para>
    /// <para>
    /// The <paramref name="nameVisibility"/> parameter determines whether or not the style name will be visible in the Excel UI.
    /// If <paramref name="nameVisibility"/> is set to <see langword="null"/>, then the style name will not be part of the resulting XLSX file.
    /// </para>
    /// <para>
    /// The style name must satisfy these requirements:
    /// <list type="bullet">
    ///   <item><description>Can not be empty or consist only of whitespace.</description></item>
    ///   <item><description>Can not start or end with whitespace.</description></item>
    ///   <item><description>Can not be longer than 255 characters.</description></item>
    ///   <item><description>Can not be equal to "Normal", as that is the name of the default font.</description></item>
    ///   <item><description>Must be unique for the spreadsheet.</description></item>
    /// </list>
    /// </para>
    /// </summary>
    public StyleId AddStyle(Style style, string name, StyleNameVisibility? nameVisibility = null)
    {
        ArgumentNullException.ThrowIfNull(style);
        ArgumentNullException.ThrowIfNull(name);

        if (string.IsNullOrWhiteSpace(name))
            ThrowHelper.NameEmptyOrWhiteSpace(nameof(name));
        if (name.Length > 255)
            ThrowHelper.StyleNameTooLong(nameof(name));
        if (char.IsWhiteSpace(name[0]) || char.IsWhiteSpace(name[^1]))
            ThrowHelper.StyleNameStartsOrEndsWithWhiteSpace(nameof(name));
        if (name.Equals("Normal", StringComparison.OrdinalIgnoreCase))
            ThrowHelper.StyleNameCanNotEqualNormal(nameof(name));
        if (nameVisibility is { } visibility)
            Guard.DefinedEnumValue(visibility);

        var styleManager = _styleManager ??= new(defaultDateTimeFormat: null);
        if (!styleManager.TryAddNamedStyle(name, style, nameVisibility, out var styleId))
            ThrowHelper.StyleNameAlreadyExists(nameof(name));

        return styleId;
    }

    /// <summary>
    /// Get the <see cref="StyleId"/> from a named style.
    /// The named style must have previously been added to the spreadsheet with <see cref="AddStyle(Style, string, StyleNameVisibility?)"/>.
    /// If the named style is not found, a <see cref="SpreadCheetahException"/> is thrown.
    /// </summary>
    public StyleId GetStyleId(string name)
    {
        ArgumentNullException.ThrowIfNull(name);

        var styleId = _styleManager?.GetStyleIdOrDefault(name);
        if (styleId is not null)
            return styleId;

        ThrowHelper.StyleNameNotFound(name);
        return null; // Unreachable
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
    /// Note that there can be max 65534 data validations in a worksheet. This method returns <see langword="false"/> if attempting to add more than that.
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
            _fileCounter.NoteForCurrentWorksheet();
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
    public void MergeCells(string cellRange)
    {
        ArgumentNullException.ThrowIfNull(cellRange);

        var cellReference = CellRangeRelativeReference.Create(cellRange);
        Worksheet.MergeCells(cellReference);
    }

    /// <summary>
    /// Embeds an image from a stream into the spreadsheet. Once an image has been embedded, it can be used
    /// in a worksheet by calling <see cref="AddImage"/> with the returned <see cref="EmbeddedImage"/> as an argument.
    /// Images can only be embedded before any worksheet has been started (i.e. before the first call to <see cref="StartWorksheetAsync"/>.
    /// Only PNG images are currently supported.
    /// </summary>
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

        var type = imageType.GetValueOrDefault();
        _fileCounter ??= new FileCounter();
        _fileCounter.AddEmbeddedImage(type);
        var embeddedImageId = _fileCounter.TotalEmbeddedImages;

        return await _zipArchiveManager.CreateImageEntryAsync(stream, buffer, type, embeddedImageId, _spreadsheetGuid, token).ConfigureAwait(false);
    }

    /// <summary>
    /// Adds an embedded image to the current worksheet. To embed an image, call <see cref="EmbedImageAsync"/>.
    /// Placement and size is defined by the <see cref="ImageCanvas"/> parameter.
    /// </summary>
    public void AddImage(ImageCanvas canvas, EmbeddedImage image, ImageOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(image);
        ImageValidator.EnsureValidCanvas(canvas, image);

        if (_spreadsheetGuid != image.SpreadsheetGuid)
            ThrowHelper.CantAddImageEmbeddedInOtherSpreadsheet();

        _fileCounter ??= new FileCounter();
        _fileCounter.TotalAddedImages++;

        var worksheetImage = new WorksheetImage(canvas, image, options?.Offset, _fileCounter.TotalAddedImages);
        Worksheet.AddImage(worksheetImage, out var firstImage);

        if (firstImage)
            _fileCounter.ImageForCurrentWorksheet();
    }

    /// <summary>
    /// Finishes the currently active table. If there are no active table, a <see cref="SpreadCheetahException"/> will be thrown.
    /// This call can be omitted if no more rows will be added to the worksheet, since finishing a worksheet will implicitly finish any active table.
    /// </summary>
    public ValueTask FinishTableAsync(CancellationToken token = default)
    {
        return Worksheet.FinishTableAsync(throwWhenNoTable: true, token);
    }

    private async ValueTask FinishAndDisposeWorksheetAsync(CancellationToken token)
    {
        if (_worksheet is not { } worksheet) return;

        await worksheet.FinishAsync(token).ConfigureAwait(false);
        await worksheet.DisposeAsync().ConfigureAwait(false);

        if (worksheet.Notes is { } notes)
        {
            var notesFileIndex = _fileCounter?.CurrentWorksheetNotesFileIndex ?? 0;
            Debug.Assert(notesFileIndex > 0);
            using var notesPooledArray = notes.ToPooledArray();

            await CommentsXml.WriteAsync(_zipArchiveManager, _buffer, notesFileIndex, notesPooledArray.Memory, token).ConfigureAwait(false);
            await VmlDrawingXml.WriteAsync(_zipArchiveManager, _buffer, notesFileIndex, notesPooledArray.Memory, token).ConfigureAwait(false);
        }

        if (worksheet.Images is { } images)
        {
            var drawingsFileIndex = _fileCounter?.CurrentWorksheetDrawingsFileIndex ?? 0;
            Debug.Assert(drawingsFileIndex > 0);
            await _zipArchiveManager.WriteDrawingAsync(_buffer, worksheet, drawingsFileIndex, images, token).ConfigureAwait(false);
            await DrawingRelsXml.WriteAsync(_zipArchiveManager, _buffer, drawingsFileIndex, images, token).ConfigureAwait(false);
        }

        if (worksheet.Tables is { } tables)
        {
            await _zipArchiveManager.WriteTablesAsync(tables, _buffer, _fileCounter, token).ConfigureAwait(false);
        }

        if (_fileCounter is { CurrentWorksheetHasRelationships: true } counter)
        {
            var worksheetIndex = _worksheets.Count;
            await WorksheetRelsXml.WriteAsync(_zipArchiveManager, _buffer, worksheetIndex, counter, token).ConfigureAwait(false);
        }

        _worksheet = null;
        _fileCounter?.ResetCurrentWorksheet();
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

        var zip = _zipArchiveManager;

        if (_documentProperties is not null)
        {
            await zip.WriteAppXmlAsync(_buffer, token).ConfigureAwait(false);
            await zip.WriteCoreXmlAsync(_documentProperties, _buffer, token).ConfigureAwait(false);
        }

        var hasStyles = _styleManager is not null;
        var includeDocumentProperties = _documentProperties is not null;

        await zip.WriteRelationshipsXmlAsync(includeDocumentProperties, _buffer, token).ConfigureAwait(false);
        await ContentTypesXml.WriteAsync(zip, _buffer, _worksheets, _fileCounter, hasStyles, includeDocumentProperties, token).ConfigureAwait(false);
        await WorkbookRelsXml.WriteAsync(zip, _buffer, _worksheets, hasStyles, token).ConfigureAwait(false);
        await WorkbookXml.WriteAsync(zip, _buffer, _worksheets, token).ConfigureAwait(false);

        if (_styleManager is not null)
            await StylesXml.WriteAsync(zip, _buffer, _styleManager, token).ConfigureAwait(false);

        _finished = true;

        // The XLSX can become corrupt if the archive is not flushed before the resulting stream is being used.
        // ZipArchive.Dispose() is currently (up to .NET 7) the only way to flush the archive to the resulting stream.
        // In case the user would use the resulting stream before Spreadsheet.Dispose(), the flush must happen here to prevent a corrupt XLSX.
        zip.Dispose();
    }

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_worksheet != null)
            await _worksheet.DisposeAsync().ConfigureAwait(false);

        _buffer.Dispose();
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _worksheet?.Dispose();
        _buffer.Dispose();
    }
}
