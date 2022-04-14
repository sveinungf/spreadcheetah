using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(RecordClassWithCustomType))]
public partial class CustomTypeContext : WorksheetRowContext
{
}
