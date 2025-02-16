using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Helpers;

internal static class TableThrowHelper
{
    [DoesNotReturn]
    public static void DuplicateHeaderName(string headerName)
        => throw new SpreadCheetahException($"Duplicate header names are not allowed in a table. Header name '{headerName}' appeared more than once.");

    [DoesNotReturn]
    public static void HeaderNameNullOrEmpty(int columnNumber)
        => throw new SpreadCheetahException(StringHelper.Invariant($"Header for table column {columnNumber} was null or empty."));

    [DoesNotReturn]
    public static void MissingHeaderName()
        => throw new SpreadCheetahException("Table must have at least one header name.");

    [DoesNotReturn]
    public static void MissingHeaderNames(int expectedCount, int actualCount)
        => throw new SpreadCheetahException(StringHelper.Invariant($"Table was expected to have {expectedCount} header names, but only {actualCount} were supplied."));

    // TODO: Move all other table related methods from ThrowHelper to this class.
}
