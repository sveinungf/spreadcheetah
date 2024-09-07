using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

[WorksheetRow(typeof(ClassWithGenericConverter))]
[WorksheetRow(typeof(ClassWithReusedConverter))]
internal partial class CellValueConverterContext : WorksheetRowContext;