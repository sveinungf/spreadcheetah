using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

internal class ClassWithGenericConverter
{
    public required string FirstName { get; init; }

    [CellValueConverter(typeof(NullToDashValueConverter<string>))]
    public required string? MiddleName { get; init; }

    public required string LastName { get; init; }

    [CellValueConverter(typeof(NullToDashValueConverter<decimal?>))]
    public required decimal? Gpa { get; init; }
}
