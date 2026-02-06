using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct BordersXmlPart(
    IList<ImmutableBorder>? borders,
    SpreadsheetBuffer buffer)
{
    private BorderXml? _currentXmlWriter;
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
            Element.Borders => TryWriteBorders(),
            _ => buffer.TryWrite("</borders>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        // The default border must come first
        var count = (borders?.Count ?? 0) + 1;
        return buffer.TryWrite(
            $"{"<borders count=\""u8}" +
            $"{count}" +
            $"{"\"><border><left/><right/><top/><bottom/><diagonal/></border>"u8}");
    }

    private bool TryWriteBorders()
    {
        if (borders is not { } bordersLocal)
            return true;

        for (; _nextIndex < bordersLocal.Count; ++_nextIndex)
        {
            var border = bordersLocal[_nextIndex];
            Debug.Assert(border != default);

            var xml = _currentXmlWriter ?? new BorderXml(border, buffer);
            if (!xml.TryWrite())
            {
                _currentXmlWriter = xml;
                return false;
            }

            _currentXmlWriter = null;
        }

        _nextIndex = 0;

        return true;
    }

    private enum Element
    {
        Header,
        Borders,
        Footer,
        Done
    }
}
