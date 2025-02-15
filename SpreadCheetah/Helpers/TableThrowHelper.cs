using SpreadCheetah.Tables;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Helpers;

internal static class TableThrowHelper
{
    [DoesNotReturn]
    public static void DuplicateHeaderName(string headerName)
        => throw new SpreadCheetahException($"Duplicate header names are not allowed in a table. Header name '{headerName}' appeared more than once.");

    [DoesNotReturn]
    public static void MustHaveHeaderNamesOrNumberOfColumnsSet()
        => throw new SpreadCheetahException($"Table must either have header names or {nameof(Table)}.{nameof(Table.NumberOfColumns)} must have a value.");

    // TODO: Move all other table related methods from ThrowHelper to this class.
}
