namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to generate code for adding the
/// specified type as a row to a worksheet.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class WorksheetRowAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="WorksheetRowAttribute"/> with the specified type.
    /// </summary>
    /// <param name="type">The type to generate source code for.</param>
#pragma warning disable CA1019 // Define accessors for attribute arguments
#pragma warning disable RCS1163 // Unused parameter.
    public WorksheetRowAttribute(Type type) { }
#pragma warning restore RCS1163 // Unused parameter.
#pragma warning restore CA1019 // Define accessors for attribute arguments
}
