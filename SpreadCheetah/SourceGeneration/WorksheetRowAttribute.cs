namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to generate code for adding the
/// specified type as a row to a worksheet.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class WorksheetRowAttribute(Type type) : Attribute;
