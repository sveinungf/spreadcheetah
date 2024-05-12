using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;

namespace SpreadCheetah.SourceGenerator.Test.Models.Combinations;

[InheritColumns]
public class ClassWithColumnAttributes(string id, string countryOfOrigin, string model, string make, int year, decimal kW, decimal length)
    : ClassWithColumnAttributesBase(id, countryOfOrigin)
{
    public string Model { get; } = model;

    [ColumnOrder(2)]
    [ColumnHeader("The make")]
    public string Make { get; } = make;

    [ColumnOrder(1)]
    public int Year { get; } = year;

#pragma warning disable IDE1006 // Naming Styles
    public decimal kW { get; } = kW;
#pragma warning restore IDE1006 // Naming Styles

    [ColumnHeader(typeof(ColumnHeaderResources), nameof(ColumnHeaderResources.Header_Length))]
    public decimal Length { get; } = length;
}

public abstract class ClassWithColumnAttributesBase(string id, string countryOfOrigin)
{
    [ColumnOrder(1000)]
    public string Id { get; } = id;

    [ColumnHeader("Country of origin")]
    public string CountryOfOrigin { get; } = countryOfOrigin;
}