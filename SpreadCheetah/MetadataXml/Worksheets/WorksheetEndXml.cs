using SpreadCheetah.CellReferences;
using SpreadCheetah.ConditionalFormatting.Internal;
using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml.Worksheets;

internal static class WorksheetEndXml
{
    public static async ValueTask WriteAsync(
        Worksheet worksheet,
        SpreadsheetBuffer buffer,
        Stream stream,
        CancellationToken token)
    {
        using var cellMerges = worksheet.CellMerges?.ToPooledArray();
        using var conditionalFormatRules = worksheet.ConditionalFormatRulesManager?.Rules.ToPooledArray();
        using var validations = worksheet.Validations?.ToPooledArray();

        var writer = new WorksheetEndXmlWriter(
            worksheet: worksheet,
            cellMerges: cellMerges?.Memory ?? default,
            conditionalFormatRules: conditionalFormatRules?.Memory ?? default,
            validations: validations?.Memory ?? default,
            buffer: buffer);

        foreach (var success in writer)
        {
            if (!success)
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
    }
}

file struct WorksheetEndXmlWriter(
    Worksheet worksheet,
    ReadOnlyMemory<CellRangeRelativeReference> cellMerges,
    ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, List<ImmutableConditionalFormatRule>>> conditionalFormatRules,
    ReadOnlyMemory<KeyValuePair<SingleCellOrCellRangeReference, DataValidation>> validations,
    SpreadsheetBuffer buffer)
    : IXmlWriter<WorksheetEndXmlWriter>
{
    private readonly int _tableCount = worksheet.Tables?.Count ?? 0;
    private ConditionalFormattingXml? _conditionalFormattingXmlWriter;
    private ConditionalFormatPriorityCounter? _conditionalFormatPriorityCounter;
    private DataValidationXml? _validationXmlWriter;
    private Element _next;
    private int _nextIndex;

    public readonly WorksheetEndXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.SheetDataEnd => buffer.TryWrite("</sheetData>"u8),
            Element.AutoFilter => TryWriteAutoFilter(),
            Element.CellMergesStart => TryWriteCellMergesStart(),
            Element.CellMerges => TryWriteCellMerges(),
            Element.CellMergesEnd => TryWriteCellMergesEnd(),
            Element.ConditionalFormatting => TryWriteConditionalFormatting(),
            Element.ValidationsStart => TryWriteValidationsStart(),
            Element.Validations => TryWriteValidations(),
            Element.ValidationsEnd => TryWriteValidationsEnd(),
            Element.Drawing => TryWriteDrawing(),
            Element.LegacyDrawing => TryWriteLegacyDrawing(),
            Element.TablePartsStart => TryWriteTablePartsStart(),
            Element.TableParts => TryWriteTableParts(),
            Element.TablePartsEnd => TryWriteTablePartsEnd(),
            _ => buffer.TryWrite("</worksheet>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteAutoFilter()
    {
        return worksheet.AutoFilterRange is not { } range
            || buffer.TryWrite($"{"<autoFilter ref=\""u8}{range}{"\"></autoFilter>"u8}");
    }

    private readonly bool TryWriteCellMergesStart()
    {
        return cellMerges.IsEmpty
            || buffer.TryWrite($"{"<mergeCells count=\""u8}{cellMerges.Length}{"\">"u8}");
    }

    private bool TryWriteCellMerges()
    {
        if (cellMerges.IsEmpty) return true;
        var span = cellMerges.Span;

        for (; _nextIndex < span.Length; ++_nextIndex)
        {
            var cellMerge = span[_nextIndex];

            var success = buffer.TryWrite($"{"<mergeCell ref=\""u8}{cellMerge.Reference}{"\"/>"u8}");
            if (!success)
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellMergesEnd()
        => cellMerges.IsEmpty || buffer.TryWrite("</mergeCells>"u8);

    private bool TryWriteConditionalFormatting()
    {
        if (conditionalFormatRules.IsEmpty) return true;
        var span = conditionalFormatRules.Span;

        for (; _nextIndex < span.Length; ++_nextIndex)
        {
            var (reference, rules) = span[_nextIndex];

            var priorityCounter = _conditionalFormatPriorityCounter ??= new();
            var writer = _conditionalFormattingXmlWriter ?? new ConditionalFormattingXml(reference, rules, priorityCounter, buffer);
            if (!writer.TryWrite())
            {
                _conditionalFormattingXmlWriter = writer;
                return false;
            }

            _conditionalFormattingXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteValidationsStart()
    {
        return validations.IsEmpty
            || buffer.TryWrite($"{"<dataValidations count=\""u8}{validations.Length}{"\">"u8}");
    }

    private bool TryWriteValidations()
    {
        if (validations.IsEmpty) return true;
        var span = validations.Span;

        for (; _nextIndex < span.Length; ++_nextIndex)
        {
            var pair = span[_nextIndex];

            var writer = _validationXmlWriter ?? new DataValidationXml(pair.Key, pair.Value, buffer);
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
        => validations.IsEmpty || buffer.TryWrite("</dataValidations>"u8);

    private readonly bool TryWriteDrawing() => worksheet.Images is null ||
        buffer.TryWrite(
            $"{"<drawing r:id=\""u8}" +
            $"{WorksheetRelationshipIds.Drawing}" +
            $"{"\"/>"u8}");

    private readonly bool TryWriteLegacyDrawing() => worksheet.Notes is null ||
        buffer.TryWrite(
            $"{"<legacyDrawing r:id=\""u8}" +
            $"{WorksheetRelationshipIds.VmlDrawing}" +
            $"{"\"/>"u8}");

    private readonly bool TryWriteTablePartsStart()
        => _tableCount == 0 || buffer.TryWrite($"{"<tableParts count=\""u8}{_tableCount}{"\">"u8}");

    private bool TryWriteTableParts()
    {
        for (; _nextIndex < _tableCount; ++_nextIndex)
        {
            var relationshipId = WorksheetRelationshipIds.TableStartId + _nextIndex;
            var success = buffer.TryWrite($"{"<tablePart r:id=\"rId"u8}{relationshipId}{"\"/>"u8}");
            if (!success)
                return false;
        }
        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteTablePartsEnd()
        => _tableCount == 0 || buffer.TryWrite("</tableParts>"u8);

    private enum Element
    {
        SheetDataEnd,
        AutoFilter,
        CellMergesStart,
        CellMerges,
        CellMergesEnd,
        ConditionalFormatting,
        ValidationsStart,
        Validations,
        ValidationsEnd,
        Drawing,
        LegacyDrawing,
        TablePartsStart,
        TableParts,
        TablePartsEnd,
        Footer,
        Done
    }
}
