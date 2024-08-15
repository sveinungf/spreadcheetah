using System.Runtime.InteropServices;

namespace SpreadCheetah.SourceGenerator.Models;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct PropertyAttributeData(
    ColumnHeader? ColumnHeader,
    ColumnOrder? ColumnOrder,
    ColumnWidth? ColumnWidth,
    CellValueTruncate? CellValueTruncate);