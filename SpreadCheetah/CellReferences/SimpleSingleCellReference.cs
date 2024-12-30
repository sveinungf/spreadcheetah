using System.Runtime.InteropServices;

namespace SpreadCheetah.CellReferences;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct SimpleSingleCellReference(ushort Column, uint Row);
