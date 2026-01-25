using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.Tables.Internal;

internal readonly record struct ImmutableTableColumnOptions(
    string? TotalRowLabel,
    TableTotalRowFunction? TotalRowFunction,
    StyleId? TotalRowStyleId)
{
    public static ImmutableTableColumnOptions From(TableColumnOptions options, StyleManager styleManager)
    {
        var totalRowStyleId = options.TotalRowStyle is { } style
            ? styleManager.AddStyleIfNotExists(style)
            : null;

        return new ImmutableTableColumnOptions(
            options.TotalRowLabel,
            options.TotalRowFunction,
            totalRowStyleId);
    }
}