using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithColumnOrdering))]
[WorksheetRow(typeof(RecordClassWithColumnOrdering))]
[WorksheetRow(typeof(StructWithColumnOrdering))]
[WorksheetRow(typeof(RecordStructWithColumnOrdering))]
[WorksheetRow(typeof(ReadOnlyStructWithColumnOrdering))]
[WorksheetRow(typeof(ReadOnlyRecordStructWithColumnOrdering))]
public partial class ColumnOrderingContext : WorksheetRowContext;