using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;

[WorksheetRow(typeof(ClassWithInheritedColumnsFirst))]
[WorksheetRow(typeof(ClassWithInheritedColumnsLast))]
[WorksheetRow(typeof(DerivedClassWithoutInheritColumns))]
[WorksheetRow(typeof(TwoLevelInheritanceClassWithInheritedColumns))]
public partial class InheritColumnsContext : WorksheetRowContext;