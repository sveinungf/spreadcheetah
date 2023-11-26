using System.Runtime.InteropServices;

namespace SpreadCheetah.Images.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct WorksheetImage(
    ImageCanvas Canvas,
    ImmutableImage Image,
    int ImageNumber);