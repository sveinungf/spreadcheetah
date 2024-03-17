using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;

[WorksheetRow(typeof(ClassWithPropertyReferenceColumnHeaders))]
[WorksheetRow(typeof(ClassWithSpecialCharacterColumnHeaders))]
public partial class ColumnHeaderContext : WorksheetRowContext;