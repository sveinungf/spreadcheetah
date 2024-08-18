using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

[WorksheetRow(typeof(ClassWithCellStyle))]
[WorksheetRow(typeof(ClassWithMultipleCellStyles))]
public partial class CellStyleContext : WorksheetRowContext;