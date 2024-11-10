using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Helpers;
using System.Collections.Immutable;

namespace SpreadCheetah.SourceGenerators;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class WorksheetRowAnalyzer : DiagnosticAnalyzer
{
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => Diagnostics.AllDescriptors;

    public override void Initialize(AnalysisContext context)
    {
        if (context is null)
            throw new ArgumentNullException(nameof(context));

        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();

        context.RegisterSymbolAction(AnalyzeContextType, SymbolKind.NamedType);
        context.RegisterSymbolAction(AnalyzeRowType, SymbolKind.NamedType);
        context.RegisterSymbolAction(AnalyzeRowTypeProperty, SymbolKind.Property);
    }

    private static void AnalyzeContextType(SymbolAnalysisContext context)
    {
        if (context.Symbol is not INamedTypeSymbol type)
            return;
        if (type is not { IsStatic: false, BaseType: { } baseType })
            return;
        if (!"SpreadCheetah.SourceGeneration.WorksheetRowContext".Equals(baseType.ToDisplayString(), StringComparison.Ordinal))
            return;
        if (type.GetAttributes() is { Length: 0 } attributes)
            return;

        var suppressWarnings = false;

        foreach (var attribute in attributes)
        {
            if (attribute.TryParseOptionsAttribute(out var options))
            {
                suppressWarnings = options.SuppressWarnings;
                break;
            }
        }

        // Right now the diagnostics below are just warnings.
        // If that changes, then this part here must also change.
        if (suppressWarnings)
            return;

        var diagnostics = new DiagnosticsReporter(context);

        foreach (var attribute in attributes)
        {
            if (!attribute.TryParseWorksheetRowAttribute(out var rowType))
                continue;

            var properties = rowType
                .GetClassAndBaseClassProperties()
                .Where(x => x.IsInstancePropertyWithPublicGetter());

            var hasProperties = false;

            foreach (var property in properties)
            {
                hasProperties = true;

                if (property.HasColumnIgnoreAttribute())
                    continue;
                if (property.HasCellValueConverterAttribute())
                    continue;
                if (property.Type.IsSupportedType())
                    continue;

                diagnostics.ReportUnsupportedPropertyType(attribute, rowType, property, context.CancellationToken);
                break;
            }

            if (!hasProperties)
                diagnostics.ReportNoPropertiesFound(attribute, rowType, context.CancellationToken);
        }
    }

    private static void AnalyzeRowType(SymbolAnalysisContext context)
    {
        if (context.Symbol is not INamedTypeSymbol type)
            return;

        var typeHasColumnOrderAttribute = type
            .GetMembers()
            .OfType<IPropertySymbol>()
            .Any(x => x.GetColumnOrderAttribute() is not null);

        // Avoid duplicate column order on base class from emitting the same error for each derived class.
        // Also some extra work is avoided for types that don't use the attribute.
        if (!typeHasColumnOrderAttribute)
            return;

        var diagnostics = new DiagnosticsReporter(context);
        var columnOrderValues = new HashSet<int>();

        foreach (var property in type.GetClassAndBaseClassProperties())
        {
            var columnOrderAttribute = property.GetColumnOrderAttribute();
            if (columnOrderAttribute is null)
                continue;

            var args = columnOrderAttribute.ConstructorArguments;
            if (args is not [{ Value: int attributeValue }])
                continue;

            if (!columnOrderValues.Add(attributeValue))
                diagnostics.ReportDuplicateColumnOrdering(columnOrderAttribute, context.CancellationToken);
        }
    }

    private static void AnalyzeRowTypeProperty(SymbolAnalysisContext context)
    {
        if (context.Symbol is not IPropertySymbol property)
            return;
        if (property.GetAttributes() is { Length: 0 })
            return;

        var diagnostics = new DiagnosticsReporter(context);
        var analyzer = new PropertyAnalyzer(diagnostics);

        analyzer.Analyze(property, context.CancellationToken);
    }
}