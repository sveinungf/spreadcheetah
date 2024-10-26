using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

[WorksheetRow(typeof(ClassWithCellCustomFormat))]
[WorksheetRow(typeof(ClassWithCellStandardFormat))]
public partial class CellFormatContext : WorksheetRowContext;