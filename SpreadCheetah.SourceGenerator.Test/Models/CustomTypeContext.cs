using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models;

[WorksheetRow(typeof(RecordWithCustomType))]
public partial class CustomTypeContext : WorksheetRowContext
{
}
