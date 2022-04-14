using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithMultipleProperties))]
[WorksheetRow(typeof(ReadOnlyStructWithMultipleProperties))]
[WorksheetRow(typeof(RecordClassWithMultipleProperties))]
[WorksheetRow(typeof(StructWithMultipleProperties))]
public partial class MultiplePropertiesContext : WorksheetRowContext
{
}
