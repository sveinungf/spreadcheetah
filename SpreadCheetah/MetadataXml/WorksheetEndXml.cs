using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetEndXml
{
    private readonly List<CellReference>? _cellMerges;
    private readonly List<KeyValuePair<CellReference, DataValidation>>? _validations;
    private readonly string? _autoFilterRange;
    private Element _next;
    private int _nextIndex;

    public WorksheetEndXml(
        List<CellReference>? cellMerges,
        List<KeyValuePair<CellReference, DataValidation>>? validations,
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

        // TODO: Data validations

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

        return true;
    }

    private readonly bool TryWriteCellMergesEnd(Span<byte> bytes, ref int bytesWritten)
        => _cellMerges is null || "</mergeCells>"u8.TryCopyTo(bytes, ref bytesWritten);

    private enum Element
    {
        SheetDataEnd,
        AutoFilter,
        CellMergesStart,
        CellMerges,
        CellMergesEnd,
        //ValidationsStart,
        //Validations,
        //ValidationsEnd,
        Footer,
        Done
    }
}
