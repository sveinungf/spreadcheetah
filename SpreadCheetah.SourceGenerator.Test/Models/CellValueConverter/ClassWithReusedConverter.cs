using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

internal class ClassWithReusedConverter
{
    [CellValueConverter(typeof(UpperCaseValueConverter))]
    public required string FirstName { get; init; }

    public required string? MiddleName { get; init; }

    [CellValueConverter(typeof(UpperCaseValueConverter))]
    public required string LastName { get; init; }

    public required decimal? Gpa { get; init; }
}
