using System.Runtime.CompilerServices;

[assembly: CLSCompliant(true)]

#if DEBUG
[assembly: InternalsVisibleTo("SpreadCheetah.Test")]
#endif