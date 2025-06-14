using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Formulas;

public class ClassWithStyledFormulas
{
    [CellStyle("Bold")]
    public Formula BoldFormula { get; set; }

    [CellFormat("#.##")]
    public Formula? FormatFormula { get; set; }
}
