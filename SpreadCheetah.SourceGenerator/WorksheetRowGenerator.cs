using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Diagnostics;
using System.Text;

namespace SpreadCheetah.SourceGenerators;

[Generator]
public class WorksheetRowGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var contextClasses = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                "SpreadCheetah.SourceGeneration.WorksheetRowAttribute",
                IsSyntaxTargetForGeneration,
                GetSemanticTargetForGeneration)
            .Where(static x => x is not null)
            .WithTrackingName(TrackingNames.Transform);

        context.RegisterSourceOutput(contextClasses, static (spc, source) => Execute(source, spc));
    }

    private static bool IsSyntaxTargetForGeneration(SyntaxNode syntaxNode, CancellationToken _) => syntaxNode is ClassDeclarationSyntax
    {
        BaseList.Types.Count: > 0
    };

    private static ContextClass? GetSemanticTargetForGeneration(GeneratorAttributeSyntaxContext context, CancellationToken token)
    {
        if (context.TargetNode is not ClassDeclarationSyntax classDeclaration)
            return null;

        if (!classDeclaration.Modifiers.Any(static x => x.IsKind(SyntaxKind.PartialKeyword)))
            return null;

        var classSymbol = context.SemanticModel.GetDeclaredSymbol(classDeclaration, token);
        if (classSymbol is not { IsStatic: false, BaseType: { } baseType })
            return null;

        if (!string.Equals("SpreadCheetah.SourceGeneration.WorksheetRowContext", baseType.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var rowTypes = new List<RowType>();

        foreach (var attribute in context.Attributes)
        {
            if (!attribute.TryParseWorksheetRowAttribute(token, out var typeSymbol, out var location))
                continue;

            var rowType = AnalyzeTypeProperties(typeSymbol, location.ToLocationInfo(), token);
            if (!rowTypes.Exists(x => string.Equals(x.FullName, rowType.FullName, StringComparison.Ordinal)))
                rowTypes.Add(rowType);
        }

        if (rowTypes.Count == 0)
            return null;

        GeneratorOptions? generatorOptions = null;

        foreach (var attribute in classSymbol.GetAttributes())
        {
            if (attribute.TryParseOptionsAttribute(out var options))
            {
                generatorOptions = options;
                break;
            }
        }

        return new ContextClass(
            DeclaredAccessibility: classSymbol.DeclaredAccessibility,
            Namespace: classSymbol.ContainingNamespace is { IsGlobalNamespace: false } ns ? ns.ToString() : null,
            Name: classSymbol.Name,
            RowTypes: rowTypes.ToEquatableArray(),
            Options: generatorOptions);
    }

    private static RowType AnalyzeTypeProperties(ITypeSymbol classType, LocationInfo? worksheetRowAttributeLocation, CancellationToken token)
    {
        var implicitOrderProperties = new List<RowTypeProperty>();
        var explicitOrderProperties = new SortedDictionary<int, RowTypeProperty>();
        var unsupportedPropertyTypeNames = new HashSet<string>(StringComparer.Ordinal);
        var diagnosticInfos = new List<DiagnosticInfo>();

        foreach (var property in GetClassAndBaseClassProperties(classType))
        {
            if (property.IsWriteOnly || property.IsStatic || property.DeclaredAccessibility != Accessibility.Public)
                continue;

            if (!property.Type.IsSupportedType())
            {
                unsupportedPropertyTypeNames.Add(property.Type.Name);
                continue;
            }

            ColumnHeader? columnHeader = null;
            ColumnOrder? columnOrder = null;
            CellValueTruncate? cellValueTruncate = null;

            foreach (var attribute in property.GetAttributes())
            {
                columnHeader ??= attribute.TryGetColumnHeaderAttribute(diagnosticInfos, token);
                columnOrder ??= attribute.TryGetColumnOrderAttribute(token);
                cellValueTruncate ??= attribute.TryGetCellValueTruncateAttribute(property.Type, diagnosticInfos, token);
            }

            var rowTypeProperty = new RowTypeProperty(property.Name, columnHeader?.ToColumnHeaderInfo(), cellValueTruncate);

            if (columnOrder is not { } order)
                implicitOrderProperties.Add(rowTypeProperty);
            else if (!explicitOrderProperties.ContainsKey(order.Value))
                explicitOrderProperties.Add(order.Value, rowTypeProperty);
            else
                diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.DuplicateColumnOrder, order.Location, new([classType.Name])));
        }

        explicitOrderProperties.AddWithImplicitKeys(implicitOrderProperties);

        return new RowType(
            DiagnosticInfos: diagnosticInfos.ToEquatableArray(),
            FullName: classType.ToString(),
            IsReferenceType: classType.IsReferenceType,
            Name: classType.Name,
            Properties: explicitOrderProperties.Values.ToEquatableArray(),
            UnsupportedPropertyTypeNames: unsupportedPropertyTypeNames.ToEquatableArray(),
            WorksheetRowAttributeLocation: worksheetRowAttributeLocation);
    }

    private static IEnumerable<IPropertySymbol> GetClassAndBaseClassProperties(ITypeSymbol? classType)
    {
        if (classType is null || string.Equals(classType.Name, "Object", StringComparison.Ordinal))
        {
            return [];
        }

        var inheritedColumnOrderStrategy = classType.GetAttributes()
            .Where(data => data.TryGetInheritedColumnOrderingAttribute().HasValue)
            .Select(data => data.TryGetInheritedColumnOrderingAttribute())
            .FirstOrDefault();

        var classProperties = classType.GetMembers().OfType<IPropertySymbol>();

        if (inheritedColumnOrderStrategy is null)
        {
            return classProperties;
        }

        var inheritedProperties = GetClassAndBaseClassProperties(classType.BaseType);

        return inheritedColumnOrderStrategy switch
        {
            InheritedColumnOrder.InheritedColumnsFirst => inheritedProperties.Concat(classProperties),
            InheritedColumnOrder.InheritedColumnsLast => classProperties.Concat(inheritedProperties),
            _ => throw new ArgumentOutOfRangeException(nameof(classType), "Unsupported inheritance strategy type")
        };
    }

    private static void Execute(ContextClass? contextClass, SourceProductionContext context)
    {
        if (contextClass is null)
            return;

        var sb = new StringBuilder();
        GenerateCode(sb, contextClass, context);

        var hintName = contextClass.Namespace is { } ns
            ? $"{ns}.{contextClass.Name}.g.cs"
            : $"{contextClass.Name}.g.cs";

        context.AddSource(hintName, sb.ToString());
    }

    private static void GenerateHeader(StringBuilder sb)
    {
        sb.AppendLine("""
            // <auto-generated />
            #nullable enable
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;
            using System;
            using System.Buffers;
            using System.Collections.Generic;
            using System.Threading;
            using System.Threading.Tasks;

            """);
    }

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass, SourceProductionContext context)
    {
        GenerateHeader(sb);

        if (contextClass.Namespace is { } ns)
            sb.Append("namespace ").AppendLine(ns);

        var accessibility = SyntaxFacts.GetText(contextClass.DeclaredAccessibility);

        sb.AppendLine($$"""
            {
                {{accessibility}} partial class {{contextClass.Name}}
                {
                    private static {{contextClass.Name}}? _default;
                    public static {{contextClass.Name}} Default => _default ??= new {{contextClass.Name}}();

                    public {{contextClass.Name}}()
                    {
                    }
            """);

        var rowTypeNames = new HashSet<string>(StringComparer.Ordinal);

        var typeIndex = 0;
        foreach (var rowType in contextClass.RowTypes)
        {
            var rowTypeName = rowType.Name;
            if (!rowTypeNames.Add(rowTypeName))
                continue;

            GenerateCodeForType(sb, typeIndex, rowType, contextClass, context);
            ++typeIndex;
        }

        sb.AppendLine("""
                }
            }
            """);
    }

    private static void GenerateCodeForType(StringBuilder sb, int typeIndex, RowType rowType,
        ContextClass contextClass, SourceProductionContext context)
    {
        ReportDiagnostics(rowType, rowType.WorksheetRowAttributeLocation, contextClass.Options, context);

        sb.AppendLine().AppendLine($$"""
                    private WorksheetRowTypeInfo<{{rowType.FullName}}>? _{{rowType.Name}};
                    public WorksheetRowTypeInfo<{{rowType.FullName}}> {{rowType.Name}} => _{{rowType.Name}}
            """);

        if (rowType.Properties.Count == 0)
        {
            sb.AppendLine($$"""
                        ??= EmptyWorksheetRowContext.CreateTypeInfo<{{rowType.FullName}}>();
            """);

            return;
        }

        sb.AppendLine(FormattableString.Invariant($$"""
                        ??= WorksheetRowMetadataServices.CreateObjectInfo<{{rowType.FullName}}>(AddHeaderRow{{typeIndex}}Async, AddAsRowAsync, AddRangeAsRowsAsync);
            """));

        GenerateAddHeaderRow(sb, typeIndex, rowType.Properties);
        GenerateAddAsRow(sb, rowType);
        GenerateAddRangeAsRows(sb, rowType);
        GenerateAddAsRowInternal(sb, rowType);
        GenerateAddRangeAsRowsInternal(sb, rowType);
        GenerateAddEnumerableAsRows(sb, rowType);
        GenerateAddCellsAsRow(sb, rowType);
    }

    private static void ReportDiagnostics(RowType rowType, LocationInfo? locationInfo, GeneratorOptions? options, SourceProductionContext context)
    {
        var suppressWarnings = options?.SuppressWarnings ?? false;

        foreach (var diagnosticInfo in rowType.DiagnosticInfos)
        {
            var isWarning = diagnosticInfo.Descriptor.DefaultSeverity == DiagnosticSeverity.Warning;
            if (isWarning && suppressWarnings)
                continue;

            context.ReportDiagnostic(diagnosticInfo.ToDiagnostic());
        }

        if (suppressWarnings) return;

        var location = locationInfo?.ToLocation();

        if (rowType.Properties.Count == 0)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.NoPropertiesFound, location, rowType.Name));

        if (rowType.UnsupportedPropertyTypeNames.FirstOrDefault() is { } unsupportedPropertyTypeName)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.UnsupportedTypeForCellValue, location, rowType.Name, unsupportedPropertyTypeName));
    }

    private static void GenerateAddHeaderRow(StringBuilder sb, int typeIndex, EquatableArray<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine().AppendLine(FormattableString.Invariant($$"""
                    private static async ValueTask AddHeaderRow{{typeIndex}}Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
                    {
                        var cells = ArrayPool<StyledCell>.Shared.Rent({{properties.Count}});
                        try
                        {
            """));

        foreach (var (i, property) in properties.Index())
        {
            var header = property.ColumnHeader?.RawString
                ?? property.ColumnHeader?.FullPropertyReference
                ?? @$"""{property.Name}""";

            sb.AppendLine(FormattableString.Invariant($"""
                            cells[{i}] = new StyledCell({header}, styleId);
            """));
        }

        sb.AppendLine($$"""
                            await spreadsheet.AddRowAsync(cells.AsMemory(0, {{properties.Count}}), token).ConfigureAwait(false);
                        }
                        finally
                        {
                            ArrayPool<StyledCell>.Shared.Return(cells, true);
                        }
                    }
            """);
    }

    private static void GenerateAddAsRow(StringBuilder sb, RowType rowType)
    {
        sb.AppendLine($$"""

                    private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, {{rowType.FullNameWithNullableAnnotation}} obj, CancellationToken token)
                    {
                        if (spreadsheet is null)
                            throw new ArgumentNullException(nameof(spreadsheet));
            """);

        if (rowType.IsReferenceType)
        {
            sb.AppendLine("""
                        if (obj is null)
                            return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            """);
        }

        sb.AppendLine("""
                        return AddAsRowInternalAsync(spreadsheet, obj, token);
                    }
            """);
    }

    private static void GenerateAddAsRowInternal(StringBuilder sb, RowType rowType)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.AppendLine($$"""

                private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, {{rowType.FullName}} obj, CancellationToken token)
                {
                    var cells = ArrayPool<DataCell>.Shared.Rent({{properties.Count}});
                    try
                    {
                        await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
                    }
                    finally
                    {
                        ArrayPool<DataCell>.Shared.Return(cells, true);
                    }
                }
        """);
    }

    private static void GenerateAddRangeAsRows(StringBuilder sb, RowType rowType)
    {
        sb.AppendLine($$"""

                private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<{{rowType.FullNameWithNullableAnnotation}}> objs, CancellationToken token)
                {
                    if (spreadsheet is null)
                        throw new ArgumentNullException(nameof(spreadsheet));
                    if (objs is null)
                        throw new ArgumentNullException(nameof(objs));
                    return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
                }
        """);
    }

    private static void GenerateAddRangeAsRowsInternal(StringBuilder sb, RowType rowType)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.Append($$"""

                private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<{{rowType.FullNameWithNullableAnnotation}}> objs, CancellationToken token)
                {
                    var cells = ArrayPool<DataCell>.Shared.Rent({{properties.Count}});
                    try
                    {
                        await AddEnumerableAsRowsAsync(spreadsheet, objs, cells, token).ConfigureAwait(false);
                    }
                    finally
                    {
                        ArrayPool<DataCell>.Shared.Return(cells, true);
                    }
                }

        """);
    }

    private static void GenerateAddEnumerableAsRows(StringBuilder sb, RowType rowType)
    {
        sb.AppendLine($$"""

                private static async ValueTask AddEnumerableAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<{{rowType.FullNameWithNullableAnnotation}}> objs, DataCell[] cells, CancellationToken token)
                {
                    foreach (var obj in objs)
                    {
                        await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
                    }
                }
        """);
    }

    private static void GenerateAddCellsAsRow(StringBuilder sb, RowType rowType)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.AppendLine($$"""

                    private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, {{rowType.FullNameWithNullableAnnotation}} obj, DataCell[] cells, CancellationToken token)
                    {
            """);

        if (rowType.IsReferenceType)
        {
            sb.AppendLine("""
                        if (obj is null)
                            return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            """);
        }

        foreach (var (i, property) in properties.Index())
        {
            if (property.CellValueTruncate?.Value is { } truncateLength)
            {
                sb.AppendLine(FormattableString.Invariant($$"""
                            var p{{i}} = obj.{{property.Name}};
                            cells[{{i}}] = p{{i}} is null || p{{i}}.Length <= {{truncateLength}} ? new DataCell(p{{i}}) : new DataCell(p{{i}}.AsMemory(0, {{truncateLength}}));
                """));
            }
            else
            {
                sb.AppendLine(FormattableString.Invariant($$"""
                            cells[{{i}}] = new DataCell(obj.{{property.Name}});
                """));
            }
        }

        sb.AppendLine($$"""
                        return spreadsheet.AddRowAsync(cells.AsMemory(0, {{properties.Count}}), token);
                    }
            """);
    }
}
