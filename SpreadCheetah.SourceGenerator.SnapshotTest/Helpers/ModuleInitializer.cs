using System.Runtime.CompilerServices;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifySourceGenerators.Enable();
    }
}
