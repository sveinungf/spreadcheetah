using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to apply a named style to the cells created for a property.
/// The named style must first be added with <see cref="Spreadsheet.AddStyle(Style, string, StyleNameVisibility?)"/>
/// before it can be used by the source generator.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellStyleAttribute(string styleName) : Attribute
{
    /// <summary>
    /// Returns the constructor's <c>styleName</c> argument value.
    /// </summary>
    public string StyleName { get; } = styleName;
}
