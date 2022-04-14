using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models;

[WorksheetRow(typeof(ClassWithProperties))]
[WorksheetRow(typeof(ReadOnlyStructWithProperties))]
[WorksheetRow(typeof(RecordWithProperties))]
[WorksheetRow(typeof(StructWithProperties))]
public partial class MultiplePropertiesContext : WorksheetRowContext
{
}
