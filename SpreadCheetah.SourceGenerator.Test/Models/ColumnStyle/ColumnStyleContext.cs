using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnStyle;

[WorksheetRow(typeof(ClassWithColumnStyle))]
public partial class ColumnStyleContext : WorksheetRowContext;