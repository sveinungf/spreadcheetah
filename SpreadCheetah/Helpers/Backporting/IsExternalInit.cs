#if NETSTANDARD2_0 || NETSTANDARD2_1
using System.ComponentModel;

namespace System.Runtime.CompilerServices;

[EditorBrowsable(EditorBrowsableState.Never)]
#pragma warning disable CA1812, CA1852 // Sealed and never instantiated
internal class IsExternalInit { }
#pragma warning restore CA1812, CA1852 // Sealed and never instantiated
#endif