using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct CellStylesXmlPart(
    List<(string, AddedStyle, StyleNameVisibility)>? namedStyles,
    SpreadsheetBuffer buffer)
{
    private string? _currentXmlEncodedName;
    private int _currentXmlEncodedNameIndex;
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
            Element.CellStyles => TryWriteCellStyles(),
            _ => buffer.TryWrite("</cellStyles>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var count = (namedStyles?.Count ?? 0) + 1;
        return buffer.TryWrite(
            $"{"<cellStyles count=\""u8}" +
            $"{count}" +
            $"{"\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"u8}");
    }

    private bool TryWriteCellStyles()
    {
        if (namedStyles is not { } namedStylesLocal)
            return true;

        for (; _nextIndex < namedStylesLocal.Count; ++_nextIndex)
        {
            var (name, _, visibility) = namedStylesLocal[_nextIndex];

            if (_currentXmlEncodedName is null)
            {
                bool? hidden = visibility == StyleNameVisibility.Hidden ? true : null;
                var hiddenAttribute = new BooleanAttribute("hidden"u8, hidden);

                if (!buffer.TryWrite(
                    $"{"<cellStyle xfId=\""u8}" +
                    $"{_nextIndex + 1}" +
                    $"{"\""u8}" +
                    $"{hiddenAttribute}" +
                    $"{" name=\""u8}"))
                {
                    return false;
                }

                _currentXmlEncodedName = name;
                _currentXmlEncodedNameIndex = 0;
            }

            if (!buffer.WriteLongString(_currentXmlEncodedName, ref _currentXmlEncodedNameIndex))
                return false;

            if (!buffer.TryWrite("\"/>"u8))
                return false;

            _currentXmlEncodedName = null;
            _currentXmlEncodedNameIndex = 0;
        }

        return true;
    }

    private enum Element
    {
        Header,
        CellStyles,
        Footer,
        Done
    }
}
