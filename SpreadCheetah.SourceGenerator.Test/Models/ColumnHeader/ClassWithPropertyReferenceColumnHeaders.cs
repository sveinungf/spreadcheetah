using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;

public class ClassWithPropertyReferenceColumnHeaders
{
    [ColumnHeader(typeof(ColumnHeaderResources), nameof(ColumnHeaderResources.Header_FirstName))]
    public string? FirstName { get; set; }

    [ColumnHeader(typeof(ColumnHeaderResources), nameof(ColumnHeaderResources.Header_LastName))]
    public string? LastName { get; set; }

    [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderNationality))]
    public string? Nationality { get; set; }

    [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAddressLine1))]
    public string? AddressLine1 { get; set; }

    [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAddressLine2))]
    public string? AddressLine2 { get; set; }

    [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAge))]
    public int Age { get; set; }
}
