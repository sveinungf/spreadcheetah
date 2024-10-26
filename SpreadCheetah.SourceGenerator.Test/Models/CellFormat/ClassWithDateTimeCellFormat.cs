using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

public class ClassWithDateTimeCellFormat
{
    public DateTime FromDate { get; set; }

    [CellFormat(StandardNumberFormat.DateAndTime)]
    public DateTime ToDate { get; set; }
}
