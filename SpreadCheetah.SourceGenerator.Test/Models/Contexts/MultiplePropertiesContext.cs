using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.MultipleProperties;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithMultipleProperties))]
[WorksheetRow(typeof(RecordClassWithMultipleProperties))]
[WorksheetRow(typeof(StructWithMultipleProperties))]
[WorksheetRow(typeof(RecordStructWithMultipleProperties))]
[WorksheetRow(typeof(ReadOnlyStructWithMultipleProperties))]
[WorksheetRow(typeof(ReadOnlyRecordStructWithMultipleProperties))]
public partial class MultiplePropertiesContext : WorksheetRowContext
{
}
