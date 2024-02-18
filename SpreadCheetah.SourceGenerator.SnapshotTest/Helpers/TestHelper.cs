using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGeneration;
using System.Collections;
using System.Collections.Immutable;
using System.Reflection;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class TestHelper
{
    private static PortableExecutableReference[] GetAssemblyReferences()
    {
        var dotNetAssemblyPath = Path.GetDirectoryName(typeof(object).Assembly.Location) ?? throw new InvalidOperationException();

        return
        [
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "mscorlib.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "netstandard.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Core.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Private.CoreLib.dll")),
            MetadataReference.CreateFromFile(Path.Combine(dotNetAssemblyPath, "System.Runtime.dll")),
            MetadataReference.CreateFromFile(typeof(WorksheetRowAttribute).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(TestHelper).Assembly.Location)
        ];
    }

    public static SettingsTask CompileAndVerify<T>(string source, bool replaceLineEndings = false, params object?[] parameters) where T : IIncrementalGenerator, new()
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);
        var references = GetAssemblyReferences();
        var compilation = CSharpCompilation.Create("Tests", [syntaxTree], references);

        var generator = new T();
#pragma warning disable S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var driver = CSharpGeneratorDriver.Create(generator);
#pragma warning restore S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var target = driver.RunGenerators(compilation);

        var settings = new VerifySettings();
        settings.UseDirectory("../Snapshots");

        if (replaceLineEndings)
            settings.ScrubLinesWithReplace(x => x.Replace("\r\n", "\n", StringComparison.Ordinal));

        var task = Verify(target, settings);

        return parameters.Length > 0
            ? task.UseParameters(parameters)
            : task;
    }

    /// <summary>
    /// Based on the implementation from:
    /// https://andrewlock.net/creating-a-source-generator-part-10-testing-your-incremental-generator-pipeline-outputs-are-cacheable/
    /// </summary>
    public static (ImmutableArray<Diagnostic> Diagnostics, string[] Output) GetGeneratedTrees<T>(
        string source,
        string[] trackingStages,
        bool assertOutputs = true)
        where T : IIncrementalGenerator, new()
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);
        var references = GetAssemblyReferences();
        var options = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary);
        var compilation = CSharpCompilation.Create("SpreadCheetah.Generated", [syntaxTree], references, options);

        // Run the generator, get the results, and assert cacheability if applicable
        var runResult = RunGeneratorAndAssertOutput<T>(compilation, trackingStages, assertOutputs);

        // Return the generator diagnostics and generated sources
        return (runResult.Diagnostics, runResult.GeneratedTrees.Select(x => x.ToString()).ToArray());
    }

    private static GeneratorDriverRunResult RunGeneratorAndAssertOutput<T>(CSharpCompilation compilation, string[] trackingNames, bool assertOutput = true)
        where T : IIncrementalGenerator, new()
    {
        var generator = new T().AsSourceGenerator();

        // ⚠ Tell the driver to track all the incremental generator outputs
        // without this, you'll have no tracked outputs!
        var opts = new GeneratorDriverOptions(
            disabledOutputs: IncrementalGeneratorOutputKind.None,
            trackIncrementalGeneratorSteps: true);

        GeneratorDriver driver = CSharpGeneratorDriver.Create([generator], driverOptions: opts);

        // Create a clone of the compilation that we will use later
        var clone = compilation.Clone();

        // Do the initial run
        // Note that we store the returned driver value, as it contains cached previous outputs
        driver = driver.RunGenerators(compilation);
        GeneratorDriverRunResult runResult = driver.GetRunResult();

        if (assertOutput)
        {
            // Run again, using the same driver, with a clone of the compilation
            var runResult2 = driver.RunGenerators(clone).GetRunResult();

            // Compare all the tracked outputs, throw if there's a failure
            AssertRunsEqual(runResult, runResult2, trackingNames);

            // verify the second run only generated cached source outputs
            var outputs = runResult2
                .Results[0]
                .TrackedOutputSteps
                .SelectMany(x => x.Value) // step executions
                .SelectMany(x => x.Outputs); // execution results

            var output = Assert.Single(outputs);
            Assert.Equal(IncrementalStepRunReason.Cached, output.Reason);
        }

        return runResult;
    }

    private static void AssertRunsEqual(
        GeneratorDriverRunResult runResult1,
        GeneratorDriverRunResult runResult2,
        string[] trackingNames)
    {
        // We're given all the tracking names, but not all the
        // stages will necessarily execute, so extract all the 
        // output steps, and filter to ones we know about
        var trackedSteps1 = GetTrackedSteps(runResult1, trackingNames);
        var trackedSteps2 = GetTrackedSteps(runResult2, trackingNames);

        // Both runs should have the same tracked steps
        var trackedSteps1Keys = trackedSteps1.Keys.ToHashSet(StringComparer.Ordinal);
        Assert.True(trackedSteps1Keys.SetEquals(trackedSteps2.Keys));

        // Get the IncrementalGeneratorRunStep collection for each run
        foreach (var (trackingName, runSteps1) in trackedSteps1)
        {
            // Assert that both runs produced the same outputs
            var runSteps2 = trackedSteps2[trackingName];
            AssertEqual(runSteps1, runSteps2, trackingName);
        }

        static Dictionary<string, ImmutableArray<IncrementalGeneratorRunStep>> GetTrackedSteps(
            GeneratorDriverRunResult runResult, string[] trackingNames)
        {
            return runResult
                .Results[0] // We're only running a single generator, so this is safe
                .TrackedSteps // Get the pipeline outputs
                .Where(step => trackingNames.Contains(step.Key, StringComparer.Ordinal))
                .ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
        }
    }

    private static void AssertEqual(
        ImmutableArray<IncrementalGeneratorRunStep> runSteps1,
        ImmutableArray<IncrementalGeneratorRunStep> runSteps2,
        string stepName)
    {
        Assert.Equal(runSteps1.Length, runSteps2.Length);

        foreach (var (runStep1, runStep2) in runSteps1.Zip(runSteps2))
        {
            // The outputs should be equal between different runs
            var outputs1 = runStep1.Outputs.Select(x => x.Value);
            var outputs2 = runStep2.Outputs.Select(x => x.Value);

            Assert.True(outputs1.SequenceEqual(outputs2), $"Step {stepName} did not produce cacheable outputs");

            // Therefore, on the second run the results should always be cached or unchanged!
            // - Unchanged is when the _input_ has changed, but the output hasn't
            // - Cached is when the the input has not changed, so the cached output is used 
            Assert.All(runStep2.Outputs, x => Assert.True(x.Reason is IncrementalStepRunReason.Cached or IncrementalStepRunReason.Unchanged));

            // Make sure we're not using anything we shouldn't
            AssertObjectGraph(runStep1);
        }
    }

    private static void AssertObjectGraph(IncrementalGeneratorRunStep runStep)
    {
        var visited = new HashSet<object>();

        // Check all of the outputs - probably overkill, but why not
        foreach (var (obj, _) in runStep.Outputs)
        {
            Visit(obj);
        }

        void Visit(object? node)
        {
            // If we've already seen this object, or it's null, stop.
            if (node is null || !visited.Add(node))
                return;

            // Make sure it's not a banned type
            Assert.IsNotAssignableFrom<Compilation>(node);
            Assert.IsNotAssignableFrom<ISymbol>(node);
            Assert.IsNotAssignableFrom<SyntaxNode>(node);

            // Examine the object
            var type = node.GetType();
            if (type.IsPrimitive || type.IsEnum || type == typeof(string))
                return;

            // If the object is a collection, check each of the values
            if (node is IEnumerable collection and not string)
            {
                foreach (object element in collection)
                {
                    // recursively check each element in the collection
                    Visit(element);
                }

                return;
            }

            // Recursively check each field in the object
            foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                var fieldValue = field.GetValue(node);
                Visit(fieldValue);
            }
        }
    }
}