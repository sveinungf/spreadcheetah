using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

[WorksheetRow(typeof(ClassWithCellCustomFormat))]
[WorksheetRow(typeof(ClassWithCellStandardFormat))]
[WorksheetRow(typeof(ClassWithMultipleCellFormats))]
public partial class CellFormatContext : WorksheetRowContext;