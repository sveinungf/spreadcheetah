using SpreadCheetah.Styling;

namespace SpreadCheetah.Helpers;

internal static class HeaderNameExtensions
{
    public static void CopyToCells(
        this ReadOnlySpan<string> headerNames,
        StyledCell[] cells,
        StyleId? styleId)
    {
        for (var i = 0; i < headerNames.Length; ++i)
        {
            cells[i] = new StyledCell(headerNames[i], styleId);
        }
    }
}
