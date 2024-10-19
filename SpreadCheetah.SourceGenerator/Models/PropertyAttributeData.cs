using System.Runtime.InteropServices;

namespace SpreadCheetah.SourceGenerator.Models;

[StructLayout(LayoutKind.Auto)]
internal record struct PropertyAttributeData
{
    public CellFormat? CellFormat { get; set; }
    public CellStyle? CellStyle { get; set; }
    public CellValueConverter? CellValueConverter { get; set; }
    public CellValueTruncate? CellValueTruncate { get; set; }
    public ColumnHeader? ColumnHeader { get; set; }
    public ColumnOrder? ColumnOrder { get; set; }
    public ColumnWidth? ColumnWidth { get; set; }
}