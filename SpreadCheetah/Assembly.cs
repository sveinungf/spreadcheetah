#if !RELEASE
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("SpreadCheetah.Benchmark")]
[assembly: InternalsVisibleTo("SpreadCheetah.Test")]
[assembly: InternalsVisibleTo("SpreadCheetah.TestHelpers")]
#endif