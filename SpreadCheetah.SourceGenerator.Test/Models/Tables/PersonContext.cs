using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Tables;

[WorksheetRow(typeof(Person))]
internal partial class PersonContext : WorksheetRowContext;