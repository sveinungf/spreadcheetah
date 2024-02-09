using FluentAssertions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGeneration;
using System.Collections;
using System.Collections.Immutable;
using System.Reflection;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

internal static class TestHelper
{
    public static SettingsTask CompileAndVerify<T>(string source, params object?[] parameters) where T : IIncrementalGenerator, new()
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

        var compilation = CSharpCompilation.Create("Tests", [syntaxTree], references);

        var generator = new T();
#pragma warning disable S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var driver = CSharpGeneratorDriver.Create(generator);
#pragma warning restore S3220 // Method calls should not resolve ambiguously to overloads with "params"
        var target = driver.RunGenerators(compilation);

        var settings = new VerifySettings();
        settings.UseDirectory("../Snapshots");

        var task = Verify(target, settings);

        return parameters.Length > 0
            ? task.UseParameters(parameters)
            : task;
    }

    public static (ImmutableArray<Diagnostic> Diagnostics, string[] Output) GetGeneratedTrees<T>(
        string[] sources, // C# source code 
        string[] stages,  // The tracking stages we expect
        bool assertOutputs = true) // You can disable cacheability checking during dev
        where T : IIncrementalGenerator, new() // T is your generator
    {
        // Convert the source files to SyntaxTrees
        IEnumerable<SyntaxTree> syntaxTrees = sources.Select(static x => CSharpSyntaxTree.ParseText(x));

        // Configure the assembly references you need
        // This will vary depending on your generator and requirements
        //var references = AppDomain.CurrentDomain.GetAssemblies()
        //    .Where(_ => !_.IsDynamic && !string.IsNullOrWhiteSpace(_.Location))
        //    .Select(_ => MetadataReference.CreateFromFile(_.Location))
        //    .Concat(new[] { MetadataReference.CreateFromFile(typeof(T).Assembly.Location) });

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

        // Create a Compilation object
        // You may want to specify other results here
        CSharpCompilation compilation = CSharpCompilation.Create(
            "SpreadCheetah.Generated",
            syntaxTrees,
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        // Run the generator, get the results, and assert cacheability if applicable
        GeneratorDriverRunResult runResult = RunGeneratorAndAssertOutput<T>(
            compilation, stages, assertOutputs);

        // Return the generator diagnostics and generated sources
        return (runResult.Diagnostics, runResult.GeneratedTrees.Select(x => x.ToString()).ToArray());
    }

    private static GeneratorDriverRunResult RunGeneratorAndAssertOutput<T>(CSharpCompilation compilation, string[] trackingNames, bool assertOutput = true)
        where T : IIncrementalGenerator, new()
    {
        ISourceGenerator generator = new T().AsSourceGenerator();

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
            GeneratorDriverRunResult runResult2 = driver
                                                  .RunGenerators(clone)
                                                  .GetRunResult();

            // Compare all the tracked outputs, throw if there's a failure
            AssertRunsEqual(runResult, runResult2, trackingNames);

            // verify the second run only generated cached source outputs
            runResult2.Results[0]
                        .TrackedOutputSteps
                        .SelectMany(x => x.Value) // step executions
                        .SelectMany(x => x.Outputs) // execution results
                        .Should()
                        .OnlyContain(x => x.Reason == IncrementalStepRunReason.Cached);
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
        trackedSteps1.Should()
                     .NotBeEmpty()
                     .And.HaveSameCount(trackedSteps2)
                     .And.ContainKeys(trackedSteps2.Keys);

        // Get the IncrementalGeneratorRunStep collection for each run
        foreach (var (trackingName, runSteps1) in trackedSteps1)
        {
            // Assert that both runs produced the same outputs
            var runSteps2 = trackedSteps2[trackingName];
            AssertEqual(runSteps1, runSteps2, trackingName);
        }

        // Local function that extracts the tracked steps
        static Dictionary<string, ImmutableArray<IncrementalGeneratorRunStep>> GetTrackedSteps(
            GeneratorDriverRunResult runResult, string[] trackingNames)
            => runResult
                    .Results[0] // We're only running a single generator, so this is safe
                    .TrackedSteps // Get the pipeline outputs
                    .Where(step => trackingNames.Contains(step.Key)) // filter to known steps
                    .ToDictionary(x => x.Key, x => x.Value); // Convert to a dictionary
    }

    private static void AssertEqual(
        ImmutableArray<IncrementalGeneratorRunStep> runSteps1,
        ImmutableArray<IncrementalGeneratorRunStep> runSteps2,
        string stepName)
    {
        runSteps1.Should().HaveSameCount(runSteps2);

        for (var i = 0; i < runSteps1.Length; i++)
        {
            var runStep1 = runSteps1[i];
            var runStep2 = runSteps2[i];

            // The outputs should be equal between different runs
            IEnumerable<object> outputs1 = runStep1.Outputs.Select(x => x.Value);
            IEnumerable<object> outputs2 = runStep2.Outputs.Select(x => x.Value);

            outputs1.Should()
                    .Equal(outputs2, $"because {stepName} should produce cacheable outputs");

            // Therefore, on the second run the results should always be cached or unchanged!
            // - Unchanged is when the _input_ has changed, but the output hasn't
            // - Cached is when the the input has not changed, so the cached output is used 
            runStep2.Outputs.Should()
                .OnlyContain(
                    x => x.Reason == IncrementalStepRunReason.Cached || x.Reason == IncrementalStepRunReason.Unchanged,
                    $"{stepName} expected to have reason {IncrementalStepRunReason.Cached} or {IncrementalStepRunReason.Unchanged}");

            // Make sure we're not using anything we shouldn't
            AssertObjectGraph(runStep1, stepName);
        }
    }

    private static void AssertObjectGraph(IncrementalGeneratorRunStep runStep, string stepName)
    {
        // Including the stepName in error messages to make it easy to isolate issues
        var because = $"{stepName} shouldn't contain banned symbols";
        var visited = new HashSet<object>();

        // Check all of the outputs - probably overkill, but why not
        foreach (var (obj, _) in runStep.Outputs)
        {
            Visit(obj);
        }

        void Visit(object node)
        {
            // If we've already seen this object, or it's null, stop.
            if (node is null || !visited.Add(node))
            {
                return;
            }

            // Make sure it's not a banned type
            node.Should()
                .NotBeOfType<Compilation>(because)
                .And.NotBeOfType<ISymbol>(because)
                .And.NotBeOfType<SyntaxNode>(because);

            // Examine the object
            Type type = node.GetType();
            if (type.IsPrimitive || type.IsEnum || type == typeof(string))
            {
                return;
            }

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
            foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                object fieldValue = field.GetValue(node);
                Visit(fieldValue);
            }
        }
    }
}