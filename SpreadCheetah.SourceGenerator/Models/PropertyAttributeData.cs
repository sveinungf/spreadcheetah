using System.Runtime.InteropServices;

namespace SpreadCheetah.SourceGenerator.Models;

[StructLayout(LayoutKind.Auto)]
internal record struct PropertyAttributeData
{
    public ColumnHeader? ColumnHeader { get; set; }
    public ColumnOrder? ColumnOrder { get; set; }
    public ColumnStyle? ColumnStyle { get; set; }
    public ColumnWidth? ColumnWidth { get; set; }
    public CellValueTruncate? CellValueTruncate { get; set; }
}