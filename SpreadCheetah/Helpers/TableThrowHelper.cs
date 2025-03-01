using SpreadCheetah.Tables;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Helpers;

internal static class TableThrowHelper
{
    [DoesNotReturn]
    public static void DuplicateHeaderName(string headerName)
        => throw new SpreadCheetahException($"Duplicate header names are not allowed in a table. Header name '{headerName}' appeared more than once.");

    [DoesNotReturn]
    public static void HeaderNameCanNotBeginOrEndWithWhitespace()
        => throw new SpreadCheetahException("Table header name can not begin or end with a whitespace.");

    [DoesNotReturn]
    public static void HeaderNameNullOrEmpty(int columnNumber)
        => throw new SpreadCheetahException(StringHelper.Invariant($"Header for table column {columnNumber} was null or empty."));

    [DoesNotReturn]
    public static void MissingHeaderName()
        => throw new SpreadCheetahException("Table must have at least one header name.");

    [DoesNotReturn]
    public static void MissingHeaderNames(int expectedCount, int actualCount)
        => throw new SpreadCheetahException(StringHelper.Invariant($"Table was expected to have {expectedCount} header names, but only {actualCount} were supplied."));

    [DoesNotReturn]
    public static void MultipleActiveTables()
        => throw new SpreadCheetahException("There are multiple active tables.");

    [DoesNotReturn]
    public static void NameAlreadyExists(string? paramName)
        => throw new ArgumentException("A table with the given name already exists.", paramName);

    [DoesNotReturn]
    public static void NameCanNotBeCorR(string? paramName)
        => throw new ArgumentException("The name can not be any of 'C', 'c', 'R', or 'r'.", paramName);

    [DoesNotReturn]
    public static void NameHasInvalidCharacters(string? paramName)
        => throw new ArgumentException("The name can only contain letters, numbers, '.', '_', and '\\'. The name must also start with a letter, '_', or '\\'.", paramName);

    [DoesNotReturn]
    public static void NameIsCellReference(string? paramName)
        => throw new ArgumentException("The name can not be a cell reference.", paramName);

    [DoesNotReturn]
    public static void NoActiveTables()
        => throw new SpreadCheetahException("There are no active tables.");

    [DoesNotReturn]
    public static void NoColumns(string tableName)
        => throw new SpreadCheetahException($"Table '{tableName}' has no columns. The number of columns is determined by calling {nameof(Spreadsheet.AddHeaderRowAsync)} for the first row of the table. Alternatively it can be set explicitly with {nameof(Table)}.{nameof(Table.NumberOfColumns)}.");

    [DoesNotReturn]
    public static void NoRows(string tableName)
        => throw new SpreadCheetahException($"Table '{tableName}' has no rows. Tables must have at least one row.");

    [DoesNotReturn]
    public static void OnlyOneActiveTableAllowed()
        => throw new SpreadCheetahException($"Currently there is support for having only one active table at a time. Another table can be added if the previous table is first finished by calling {nameof(Spreadsheet.FinishTableAsync)}.");
}
