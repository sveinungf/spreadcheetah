namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

public static class ColumnHeaders
{
    public static string HeaderNationality { get; } = "The nationality";
    public static string HeaderAddressLine1 => "Address line 1";
    public static string? HeaderAddressLine2 => null;
    public static string? HeaderAge => $"Age (in {DateTime.UtcNow.Year})";
}
