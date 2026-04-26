using SpreadCheetah.CellReferences;
using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.MetadataXml.Worksheets;

internal struct ConditionalFormattingXml(
    SingleCellOrCellRangeReference reference,
    List<ImmutableConditionalFormatRule> rules,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;

#pragma warning disable EPS12 // A struct member can be made readonly
    public bool TryWrite()
#pragma warning restore EPS12 // A struct member can be made readonly
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Rules => TryWriteRules(),
            _ => buffer.TryWrite("</conditionalFormatting>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return buffer.TryWrite(
            $"{"<conditionalFormatting sqref=\""u8}" +
            $"{reference.Reference}" +
            $"{"\">"u8}");
    }

    private bool TryWriteRules()
    {
        for (; _nextIndex < rules.Count; ++_nextIndex)
        {
            var rule = rules[_nextIndex];
            if (!rule.TryWrite(buffer, _nextIndex + 1))
                return false;
        }

        return true;
    }

    private enum Element
    {
        Header,
        Rules,
        Footer,
        Done
    }
}
