using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

public class ClassWithMultipleCellFormats
{
    [CellFormat("#.00")]
    public int Id { get; set; }

    [CellFormat(StandardNumberFormat.LongDate)]
    public DateTime FromDate { get; set; }

    [CellFormat("#.00")]
    public decimal Price { get; set; }
}
