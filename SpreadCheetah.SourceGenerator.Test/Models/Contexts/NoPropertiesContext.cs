using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithNoProperties))]
[WorksheetRow(typeof(ReadOnlyStructWithNoProperties))]
[WorksheetRow(typeof(RecordClassWithNoProperties))]
[WorksheetRow(typeof(StructWithNoProperties))]
public partial class NoPropertiesContext : WorksheetRowContext
{
}
