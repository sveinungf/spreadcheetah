using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.DefaultColumnWidth;

[WorksheetRow(typeof(ClassWithDefaultColumnWidth))]
public partial class DefaultColumnWidthContext : WorksheetRowContext;