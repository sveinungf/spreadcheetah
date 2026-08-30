using System.Runtime.InteropServices;

namespace SpreadCheetah.Helpers;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct BufferWriteProgress(int Step, int Index);
