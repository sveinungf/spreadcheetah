using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Helpers;
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

        context.RegisterSymbolAction(AnalyzeNamedType, SymbolKind.NamedType);
        context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
    }

    private static void AnalyzeNamedType(SymbolAnalysisContext context)
    {
        if (context.Symbol is not INamedTypeSymbol type)
            return;
        if (type.GetMembers() is { Length: 0 } members)
            return;

        var typeHasColumnOrderAttribute = members
            .OfType<IPropertySymbol>()
            .Any(x => x.GetAttributes().Any(a => a.IsColumnOrder()));

        // Avoid duplicate column order on base class from emitting the same error for each derived class.
        // Also some extra work is avoided for types that don't use the attribute.
        if (!typeHasColumnOrderAttribute)
            return;

        var diagnosticsReporter = new DiagnosticsReporter(context);
        var columnOrderValues = new HashSet<int>();

        foreach (var property in type.GetClassAndBaseClassProperties())
        {
            var columnOrderAttribute = property
                .GetAttributes()
                .FirstOrDefault(x => x.IsColumnOrder());

            if (columnOrderAttribute is null)
                continue;

            var args = columnOrderAttribute.ConstructorArguments;
            if (args is not [{ Value: int attributeValue }])
                continue;

            if (!columnOrderValues.Add(attributeValue))
                diagnosticsReporter.ReportDuplicateColumnOrdering(columnOrderAttribute, context.CancellationToken);
        }
    }

    private static void AnalyzeProperty(SymbolAnalysisContext context)
    {
        if (context.Symbol is not IPropertySymbol property)
            return;
        if (property.GetAttributes() is { Length: 0 } attributes)
            return;

        var diagnosticsReporter = new DiagnosticsReporter(context);
        var analyzer = new PropertyAnalyzer(diagnosticsReporter);

        var data = analyzer.Analyze(property, context.CancellationToken);

        if (data is { CellValueConverter: not null, CellValueTruncate: not null })
        {
            var cellValueTruncateAttribute = attributes
                .First(x => Attributes.CellValueTruncate.Equals(x.AttributeClass?.ToDisplayString(), StringComparison.Ordinal));

            diagnosticsReporter.ReportAttributeCombinationNotSupported(cellValueTruncateAttribute,
                "CellValueConverterAttribute", context.CancellationToken);
        }
    }
}