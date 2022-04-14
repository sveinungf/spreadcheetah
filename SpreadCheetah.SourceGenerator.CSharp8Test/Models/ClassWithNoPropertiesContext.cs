using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    [WorksheetRow(typeof(ClassWithNoProperties))]
    [WorksheetRowGenerationOptions(SuppressWarnings = true)]
    public partial class ClassWithNoPropertiesContext : WorksheetRowContext
    {
    }
}
