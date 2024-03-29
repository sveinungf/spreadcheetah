using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;

namespace SpreadCheetah.SourceGenerator.Test.Models.Combinations;

public class ClassWithColumnAttributes(string model, string make, int year, decimal kW, decimal length)
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
