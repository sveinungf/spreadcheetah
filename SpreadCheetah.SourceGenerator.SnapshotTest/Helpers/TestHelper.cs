using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class TestHelper
{
    public static SettingsTask CompileAndVerify<T>(string source) where T : IIncrementalGenerator, new()
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        var dotNetAssemblyPath = Path.GetDirectoryName(typeof(object).Assembly.Location) ?? throw new InvalidOperationException();

        var references = new[]
        {
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "mscorlib.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "netstandard.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Core.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Private.CoreLib.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Runtime.dll")),
            MetadataReference.CreateFromFile(typeof(WorksheetRowAttribute).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(TestHelper).Assembly.Location)
        };

        var compilation = CSharpCompilation.Create("Tests", new[] { syntaxTree }, references);

        var generator = new T();
#pragma warning disable S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var driver = CSharpGeneratorDriver.Create(generator);
#pragma warning restore S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var target = driver.RunGenerators(compilation);

        var settings = new VerifySettings();
        settings.UseDirectory("../Snapshots");

        return Verify(target, settings);
    }
}