using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    [WorksheetRow(typeof(ClassWithNoProperties))]
    public partial class ClassWithNoPropertiesContext : WorksheetRowContext
    {
    }
}
