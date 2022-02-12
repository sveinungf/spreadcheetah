using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class TestHelper
{
    public static SettingsTask CompileAndVerify(string source)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        var references = new[]
        {
            MetadataReference.CreateFromFile(typeof(TestHelper).Assembly.Location)
        };

        var compilation = CSharpCompilation.Create("Tests", new[] { syntaxTree }, references);

        var generator = new RowCellsGenerator();
        var driver = CSharpGeneratorDriver.Create(generator);
        var target = driver.RunGenerators(compilation);

        var settings = new VerifySettings();
        settings.UseDirectory("../Snapshots");

        return Verify(target, settings);
    }
}