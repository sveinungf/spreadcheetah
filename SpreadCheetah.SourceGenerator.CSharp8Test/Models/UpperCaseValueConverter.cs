using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    public class UpperCaseValueConverter : CellValueConverter<string>
    {
        public override DataCell ConvertToDataCell(string value)
        {
            return new DataCell(value.ToUpperInvariant());
        }
    }
}
