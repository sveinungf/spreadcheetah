using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

[WorksheetRow(typeof(ClassWithCellStyle))]
[WorksheetRow(typeof(ClassWithCellStyleOnDateTimeProperty))]
[WorksheetRow(typeof(ClassWithCellStyleOnTruncatedProperty))]
[WorksheetRow(typeof(ClassWithMultipleCellStyles))]
[WorksheetRow(typeof(RecordWithCellStyle))]
public partial class CellStyleContext : WorksheetRowContext;