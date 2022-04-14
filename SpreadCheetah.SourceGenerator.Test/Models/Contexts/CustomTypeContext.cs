using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(RecordClassWithCustomType))]
[WorksheetRowGenerationOptions(SuppressWarnings = true)]
public partial class CustomTypeContext : WorksheetRowContext
{
}
