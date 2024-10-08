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
    [CellValueTruncate(8)]
    public string Make { get; } = make;

    [ColumnOrder(1)]
    [ColumnWidth(10)]
    [CellStyle("Year style")]
    [CellValueConverter(typeof(LeapYearValueConverter))]
    public int Year { get; } = year;

#pragma warning disable IDE1006 // Naming Styles
    public decimal kW { get; } = kW;
#pragma warning restore IDE1006 // Naming Styles

    [ColumnHeader(typeof(ColumnHeaderResources), nameof(ColumnHeaderResources.Header_Length))]
    [CellValueConverter(typeof(LengthInCmValueConverter))]
    public decimal Length { get; } = length;
}

public abstract class ClassWithColumnAttributesBase(string id, string countryOfOrigin)
{
    [ColumnOrder(1000)]
    public string Id { get; } = id;

    [ColumnHeader("Country of origin")]
    [ColumnWidth(40)]
    public string CountryOfOrigin { get; } = countryOfOrigin;
}