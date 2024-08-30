using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

[WorksheetRow(typeof(ClassWithInvalidCellStyleName))]
public partial class InvalidCellStyleContext : WorksheetRowContext;