using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    [WorksheetRow(typeof(ClassWithMultipleProperties))]
    public partial class ClassWithMultiplePropertiesContext : WorksheetRowContext
    {
    }
}
