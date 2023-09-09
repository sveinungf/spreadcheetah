using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetEndXml : IXmlWriter
{
    private readonly List<CellRangeRelativeReference>? _cellMerges;
    private readonly ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>> _validations;
    private readonly string? _autoFilterRange;
    private readonly bool _hasNotes;
    private DataValidationXml? _validationXmlWriter;
    private Element _next;
    private int _nextIndex;

    public WorksheetEndXml(
        List<CellRangeRelativeReference>? cellMerges,
        ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>? validations,
        string? autoFilterRange,
        bool hasNotes)
    {
        _cellMerges = cellMerges;
        _validations = validations ?? ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>.Empty;
        _autoFilterRange = autoFilterRange;
        _hasNotes = hasNotes;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.SheetDataEnd && !Advance("</sheetData>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.AutoFilter && !Advance(TryWriteAutoFilter(bytes, ref bytesWritten))) return false;
        if (_next == Element.CellMergesStart && !Advance(TryWriteCellMergesStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.CellMerges && !Advance(TryWriteCellMerges(bytes, ref bytesWritten))) return false;
        if (_next == Element.CellMergesEnd && !Advance(TryWriteCellMergesEnd(bytes, ref bytesWritten))) return false;
        if (_next == Element.ValidationsStart && !Advance(TryWriteValidationsStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.Validations && !Advance(TryWriteValidations(bytes, ref bytesWritten))) return false;
        if (_next == Element.ValidationsEnd && !Advance(TryWriteValidationsEnd(bytes, ref bytesWritten))) return false;
        if (_next == Element.LegacyDrawing && !Advance(TryWriteLegacyDrawing(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance("</worksheet>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteAutoFilter(Span<byte> bytes, ref int bytesWritten)
    {
        if (_autoFilterRange is not { } range) return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<autoFilter ref=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(range, span, ref written)) return false;
        if (!"\"></autoFilter>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private readonly bool TryWriteCellMergesStart(Span<byte> bytes, ref int bytesWritten)
    {
        if (_cellMerges is not { } cellMerges) return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<mergeCells count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(cellMerges.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteCellMerges(Span<byte> bytes, ref int bytesWritten)
    {
        if (_cellMerges is not { } cellMerges) return true;

        for (; _nextIndex < cellMerges.Count; ++_nextIndex)
        {
            var cellMerge = cellMerges[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!"<mergeCell ref=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(cellMerge.Reference, span, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellMergesEnd(Span<byte> bytes, ref int bytesWritten)
        => _cellMerges is null || "</mergeCells>"u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteValidationsStart(Span<byte> bytes, ref int bytesWritten)
    {
        if (_validations.IsEmpty) return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<dataValidations count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_validations.Length, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteValidations(Span<byte> bytes, ref int bytesWritten)
    {
        if (_validations.IsEmpty) return true;
        var validations = _validations.Span;

        for (; _nextIndex < validations.Length; ++_nextIndex)
        {
            var pair = validations[_nextIndex];

            var writer = _validationXmlWriter ?? new DataValidationXml(pair.Key, pair.Value);
            if (!writer.TryWrite(bytes, ref bytesWritten))
            {
                _validationXmlWriter = writer;
                return false;
            }

            _validationXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteValidationsEnd(Span<byte> bytes, ref int bytesWritten)
        => _validations.IsEmpty || "</dataValidations>"u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteLegacyDrawing(Span<byte> bytes, ref int bytesWritten)
        => !_hasNotes || """<legacyDrawing r:id="rId1"/>"""u8.TryCopyTo(bytes, ref bytesWritten);

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
        LegacyDrawing,
        Footer,
        Done
    }
}
