using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

[WorksheetRow(typeof(ClassWithCellValueConverter))]
[WorksheetRow(typeof(ClassWithCellValueConverterAndCellStyle))]
[WorksheetRow(typeof(ClassWithGenericConverter))]
[WorksheetRow(typeof(ClassWithReusedConverter))]
[WorksheetRow(typeof(ClassWithCellValueConverterOnCustomType))]
internal partial class CellValueConverterContext : WorksheetRowContext;