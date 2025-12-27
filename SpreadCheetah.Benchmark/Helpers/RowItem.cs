namespace SpreadCheetah.Benchmark.Helpers;

public record RowItem
{
    public required int A { get; init; }
    public required string B { get; init; }
    public required string? C { get; init; }
    public required int? D { get; init; }
    public required double E { get; init; }
    public required bool F { get; init; }
    public required string G { get; init; }
    public required bool H { get; init; }
    public required double? I { get; init; }
    public required int J { get; init; }
    public required DateTime? K { get; init; }
    public required long L { get; init; }
}