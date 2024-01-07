using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.Accessibility;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(InternalAccessibilityClassWithSingleProperty))]
internal partial class InternalAccessibilityContext : WorksheetRowContext
{
}
