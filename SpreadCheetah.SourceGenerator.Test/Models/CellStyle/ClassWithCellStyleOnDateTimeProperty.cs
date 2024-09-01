using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithCellStyleOnDateTimeProperty
{
    [CellStyle("Created style")]
    public DateTime CreatedDate { get; set; }
}
