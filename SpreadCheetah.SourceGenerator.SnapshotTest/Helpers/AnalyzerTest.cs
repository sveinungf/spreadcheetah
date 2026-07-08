using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Testing;
using Microsoft.CodeAnalysis.Testing;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class AnalyzerTest
{
    public static CSharpAnalyzerTest<WorksheetRowAnalyzer, DefaultVerifier> CreateContext(
        LanguageVersion? languageVersion = null)
    {
        var result = new CSharpAnalyzerTest<WorksheetRowAnalyzer, DefaultVerifier>
        {
            TestState =
            {
                ReferenceAssemblies = ReferenceAssemblies.Net.Net100,
                AdditionalReferences =
                {
                    MetadataReference.CreateFromFile(typeof(WorksheetRowAttribute).Assembly.Location)
                }
            }
        };

        if (languageVersion is { } version)
        {
            result.SolutionTransforms.Add((solution, projectId) =>
            {
                var newOptions = new CSharpParseOptions(version);
                return solution.WithProjectParseOptions(projectId, newOptions);
            });
        }

        return result;
    }
}
