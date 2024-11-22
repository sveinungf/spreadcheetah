using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.Accessibility;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(DefaultAccessibilityClassWithSingleProperty))]
#pragma warning disable IDE0040 // Add accessibility modifiers
partial class DefaultAccessibilityContext : WorksheetRowContext
#pragma warning restore IDE0040 // Add accessibility modifiers
{
}
