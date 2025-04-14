namespace SpreadCheetah.Metadata;

/// <summary>
/// Provides document properties on the resulting XLSX file.
/// </summary>
public sealed record DocumentProperties
{
    /// <summary>
    /// The default instance. It is internal because it should not be mutated.
    /// </summary>
    internal static readonly DocumentProperties Default = new();

    /// <summary>Author of the spreadsheet.</summary>
    public string? Author { get; set; }

    /// <summary>Subject of the spreadsheet.</summary>
    public string? Subject { get; set; }

    /// <summary>Title of the spreadsheet.</summary>
    public string? Title { get; set; }
}
