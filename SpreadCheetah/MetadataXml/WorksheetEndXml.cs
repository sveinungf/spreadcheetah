using SpreadCheetah.CellReferences;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetEndXml
{
    private readonly ReadOnlyMemory<CellRangeRelativeReference> _cellMerges;
    private readonly ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>> _validations;
    private readonly string? _autoFilterRange;
    private readonly bool _hasImages;
    private readonly bool _hasNotes;
    private readonly SpreadsheetBuffer _buffer;
    private DataValidationXml? _validationXmlWriter;
    private Element _next;
    private int _nextIndex;

    private WorksheetEndXml(
        ReadOnlyMemory<CellRangeRelativeReference> cellMerges,
        ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>> validations,
        string? autoFilterRange,
        bool hasNotes,
        bool hasImages,
        SpreadsheetBuffer buffer)
    {
        _cellMerges = cellMerges;
        _validations = validations;
        _autoFilterRange = autoFilterRange;
        _hasImages = hasImages;
        _hasNotes = hasNotes;
        _buffer = buffer;
    }

    public static async ValueTask WriteAsync(
        ReadOnlyMemory<CellRangeRelativeReference>? cellMerges,
        ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>? validations,
        string? autoFilterRange,
        bool hasNotes,
        bool hasImages,
        SpreadsheetBuffer buffer,
        Stream stream,
        CancellationToken token)
    {
        var writer = new WorksheetEndXml(
            cellMerges: cellMerges ?? ReadOnlyMemory<CellRangeRelativeReference>.Empty,
            validations: validations ?? ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>.Empty,
            autoFilterRange: autoFilterRange,
            hasNotes: hasNotes,
            hasImages: hasImages,
            buffer: buffer);

        foreach (var success in writer)
        {
            if (!success)
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
    }

    public readonly WorksheetEndXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.SheetDataEnd => _buffer.TryWrite("</sheetData>"u8),
            Element.AutoFilter => TryWriteAutoFilter(),
            Element.CellMergesStart => TryWriteCellMergesStart(),
            Element.CellMerges => TryWriteCellMerges(),
            Element.CellMergesEnd => TryWriteCellMergesEnd(),
            Element.ValidationsStart => TryWriteValidationsStart(),
            Element.Validations => TryWriteValidations(),
            Element.ValidationsEnd => TryWriteValidationsEnd(),
            Element.Drawing => TryWriteDrawing(),
            Element.LegacyDrawing => TryWriteLegacyDrawing(),
            // TODO: Write "tableParts"
            _ => _buffer.TryWrite("</worksheet>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteAutoFilter()
    {
        return _autoFilterRange is not { } range
            || _buffer.TryWrite($"{"<autoFilter ref=\""u8}{range}{"\"></autoFilter>"u8}");
    }

    private readonly bool TryWriteCellMergesStart()
    {
        return _cellMerges.IsEmpty
            || _buffer.TryWrite($"{"<mergeCells count=\""u8}{_cellMerges.Length}{"\">"u8}");
    }

    private bool TryWriteCellMerges()
    {
        if (_cellMerges.IsEmpty) return true;
        var cellMerges = _cellMerges.Span;

        for (; _nextIndex < cellMerges.Length; ++_nextIndex)
        {
            var cellMerge = cellMerges[_nextIndex];

            var success = _buffer.TryWrite($"{"<mergeCell ref=\""u8}{cellMerge.Reference}{"\"/>"u8}");
            if (!success)
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellMergesEnd()
        => _cellMerges.IsEmpty || _buffer.TryWrite("</mergeCells>"u8);

    private readonly bool TryWriteValidationsStart()
    {
        return _validations.IsEmpty
            || _buffer.TryWrite($"{"<dataValidations count=\""u8}{_validations.Length}{"\">"u8}");
    }

    private bool TryWriteValidations()
    {
        if (_validations.IsEmpty) return true;
        var validations = _validations.Span;

        for (; _nextIndex < validations.Length; ++_nextIndex)
        {
            var pair = validations[_nextIndex];

            var writer = _validationXmlWriter ?? new DataValidationXml(pair.Key, pair.Value, _buffer);
            if (!writer.TryWrite())
            {
                _validationXmlWriter = writer;
                return false;
            }

            _validationXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteValidationsEnd()
        => _validations.IsEmpty || _buffer.TryWrite("</dataValidations>"u8);

    private readonly bool TryWriteDrawing() => !_hasImages ||
        _buffer.TryWrite(
            $"{"<drawing r:id=\""u8}" +
            $"{WorksheetRelationshipIds.Drawing}" +
            $"{"\"/>"u8}");

    private readonly bool TryWriteLegacyDrawing()=> !_hasNotes ||
        _buffer.TryWrite(
            $"{"<legacyDrawing r:id=\""u8}" +
            $"{WorksheetRelationshipIds.VmlDrawing}" +
            $"{"\"/>"u8}");

    private enum Element
    {
        SheetDataEnd,
        AutoFilter,
        CellMergesStart,
        CellMerges,
        CellMergesEnd,
        ValidationsStart,
        Validations,
        ValidationsEnd,
        Drawing,
        LegacyDrawing,
        Footer,
        Done
    }
}
