using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Testing;
using Microsoft.CodeAnalysis.Testing;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class AnalyzerTest
{
    public static CSharpAnalyzerTest<WorksheetRowAnalyzer, DefaultVerifier> CreateContext()
    {
        return new CSharpAnalyzerTest<WorksheetRowAnalyzer, DefaultVerifier>
        {
            TestState =
            {
                ReferenceAssemblies = ReferenceAssemblies.Net.Net90,
                AdditionalReferences =
                {
                    MetadataReference.CreateFromFile(typeof(WorksheetRowAttribute).Assembly.Location)
                }
            }
        };
    }
}
