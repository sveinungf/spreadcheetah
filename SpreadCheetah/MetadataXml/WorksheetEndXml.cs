using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetEndXml : IXmlWriter
{
    private readonly List<CellRangeRelativeReference>? _cellMerges;
    private readonly List<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>? _validations;
    private readonly string? _autoFilterRange;
    private DataValidationXml? _validationXmlWriter;
    private Element _next;
    private int _nextIndex;

    public WorksheetEndXml(
        List<CellRangeRelativeReference>? cellMerges,
        List<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>>? validations,
        string? autoFilterRange)
    {
        _cellMerges = cellMerges;
        _validations = validations;
        _autoFilterRange = autoFilterRange;
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
        if (_validations is not { } validations) return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<dataValidations count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(validations.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteValidations(Span<byte> bytes, ref int bytesWritten)
    {
        if (_validations is not { } validations) return true;

        for (; _nextIndex < validations.Count; ++_nextIndex)
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
        => _validations is null || "</dataValidations>"u8.TryCopyTo(bytes, ref bytesWritten);

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
        Footer,
        Done
    }
}
