using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Used by code generated by the source generator and is not intended to be used directly.
/// This class is used to cache dependencies to avoid redundant lookups.
/// The class is immutable and the cache will return the same instance once it has been created.
/// </summary>
public sealed record WorksheetRowDependencyInfo
{
    /// <summary>
    /// Used by code generated by the source generator and is not intended to be used directly.
    /// A list of style IDs based on styles defined by attributes.
    /// </summary>
    public IReadOnlyList<StyleId> StyleIds { get; init; } = [];
}