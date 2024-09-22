using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Models;
using System.Collections.Immutable;

namespace SpreadCheetah.SourceGenerators;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public class WorksheetRowAnalyzer : DiagnosticAnalyzer
{
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => Diagnostics.AllDescriptors;

    public override void Initialize(AnalysisContext context)
    {
        if (context is null)
            throw new ArgumentNullException(nameof(context));

        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();

        context.RegisterSymbolAction(AnalyzeNamedType, SymbolKind.Property);
    }

    private static void AnalyzeNamedType(SymbolAnalysisContext context)
    {
        if (context.Symbol is not IPropertySymbol property)
            return;

        var attributes = property.GetAttributes();
        if (attributes.Length == 0)
            return;

        var diagnosticsReporter = new DiagnosticsReporter(context);
        var analyzer = new PropertyAnalyzer(diagnosticsReporter);

        analyzer.Analyze(property, context.CancellationToken);
    }
}