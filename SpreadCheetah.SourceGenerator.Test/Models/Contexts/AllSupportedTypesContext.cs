using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithAllSupportedTypes))]
public partial class AllSupportedTypesContext : WorksheetRowContext
{
}
