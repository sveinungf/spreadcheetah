using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

public class ClassWithCellStandardFormat
{
    [CellFormat(StandardNumberFormat.TwoDecimalPlaces)]
    public decimal Price { get; set; }
}
