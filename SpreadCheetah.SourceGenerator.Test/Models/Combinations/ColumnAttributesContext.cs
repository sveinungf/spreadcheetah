using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Combinations;

[WorksheetRow(typeof(ClassWithColumnAttributes))]
public partial class ColumnAttributesContext : WorksheetRowContext;