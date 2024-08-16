using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

[WorksheetRow(typeof(ClassWithCellStyle))]
public partial class CellStyleContext : WorksheetRowContext;