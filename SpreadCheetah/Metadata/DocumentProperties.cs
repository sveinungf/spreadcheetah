namespace SpreadCheetah.Metadata;

public sealed record DocumentProperties
{
    /// <summary>
    /// The default instance. It is internal because it should not be mutated.
    /// </summary>
    internal static readonly DocumentProperties Default = new();

    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Title { get; set; }
}
