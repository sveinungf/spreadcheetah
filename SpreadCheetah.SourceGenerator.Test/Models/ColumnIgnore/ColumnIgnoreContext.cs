using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnIgnore;

[WorksheetRow(typeof(ClassWithMultipleProperties))]
public partial class ColumnIgnoreContext : WorksheetRowContext;