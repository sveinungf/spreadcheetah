using SpreadCheetah.CellReferences;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Images.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct WorksheetImage(
    SingleCellRelativeReference Reference,
    ImmutableImage Image);