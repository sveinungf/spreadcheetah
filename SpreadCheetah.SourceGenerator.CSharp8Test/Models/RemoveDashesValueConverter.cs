using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    internal class RemoveDashesValueConverter : CellValueConverter<string>
    {
        public override DataCell ConvertToDataCell(string value)
        {
            return new DataCell(value.Replace("-", ""));
        }
    }
}
