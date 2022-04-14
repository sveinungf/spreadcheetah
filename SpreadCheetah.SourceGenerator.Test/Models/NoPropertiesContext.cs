using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models;

[WorksheetRow(typeof(ClassWithNoProperties))]
[WorksheetRow(typeof(ReadOnlyStructWithNoProperties))]
[WorksheetRow(typeof(RecordWithNoProperties))]
[WorksheetRow(typeof(StructWithNoProperties))]
public partial class NoPropertiesContext : WorksheetRowContext
{
}
