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
        var propertiesWithStyleAttributes = 0;

        foreach (var property in GetClassAndBaseClassProperties(classType))
        {
            if (property.IsWriteOnly || property.IsStatic || property.DeclaredAccessibility != Accessibility.Public)
                continue;

            if (!property.Type.IsSupportedType())
            {
                unsupportedPropertyTypeNames.Add(property.Type.Name);
                continue;
            }

            var data = property
                .GetAttributes()
                .MapToPropertyAttributeData(property.Type, diagnosticInfos, token);

            if (data.CellStyle is not null)
                propertiesWithStyleAttributes++;

            var rowTypeProperty = new RowTypeProperty(
                Name: property.Name,
                ColumnHeader: data.ColumnHeader?.ToColumnHeaderInfo(),
                CellStyle: data.CellStyle,
                ColumnWidth: data.ColumnWidth,
                CellValueTruncate: data.CellValueTruncate,
                PropertyCellValueMapper: data.CellValueMapper);

            if (data.ColumnOrder is not { } order)
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
            PropertiesWithStyleAttributes: propertiesWithStyleAttributes,
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
            using SpreadCheetah.Styling;
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

        GenerateConstructTruncatedCells(sb);

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

        var properties = rowType.Properties;
        var doGenerateCreateWorksheetOptions = properties.Any(static x => x.ColumnWidth is not null);
        var doGenerateCreateWorksheetRowDependencyInfo = properties.Any(static x => x.CellStyle is not null);

        sb.Append(FormattableString.Invariant($$"""
                        ??= WorksheetRowMetadataServices.CreateObjectInfo<{{rowType.FullName}}>(
                            AddHeaderRow{{typeIndex}}Async, AddAsRowAsync, AddRangeAsRowsAsync
            """));

        if (doGenerateCreateWorksheetOptions)
            sb.Append(", CreateWorksheetOptions").Append(typeIndex);
        else
            sb.Append(", null");

        if (doGenerateCreateWorksheetRowDependencyInfo)
            sb.Append(", CreateWorksheetRowDependencyInfo").Append(typeIndex);

        sb.AppendLine(");");

        if (doGenerateCreateWorksheetOptions)
            GenerateCreateWorksheetOptions(sb, typeIndex, properties);

        var cellStyleToStyleIdIndex = doGenerateCreateWorksheetRowDependencyInfo
            ? GenerateCreateWorksheetRowDependencyInfo(sb, typeIndex, properties)
            : [];

        GenerateCellValueMappersArray(sb, rowType);
        GenerateAddHeaderRow(sb, typeIndex, properties);
        GenerateAddAsRow(sb, rowType);
        GenerateAddRangeAsRows(sb, rowType);
        GenerateAddAsRowInternal(sb, rowType);
        GenerateAddRangeAsRowsInternal(sb, rowType);
        GenerateAddCellsAsRow(sb, rowType, cellStyleToStyleIdIndex);
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

    private static void GenerateCreateWorksheetOptions(StringBuilder sb, int typeIndex, EquatableArray<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Any(static x => x.ColumnWidth is not null));

        sb.AppendLine(FormattableString.Invariant($$"""

                    private static SpreadCheetah.Worksheets.WorksheetOptions CreateWorksheetOptions{{typeIndex}}()
                    {
                        var options = new SpreadCheetah.Worksheets.WorksheetOptions();
            """));

        foreach (var (i, property) in properties.Index())
        {
            if (property.ColumnWidth is not { } columnWidth)
                continue;

            var width = columnWidth.Width;
            var columnNumber = i + 1;

            sb.AppendLine(FormattableString.Invariant($"""
                        options.Column({columnNumber}).Width = {width};
            """));
        }

        sb.AppendLine("""
                        return options;
                    }
            """);
    }

    private static Dictionary<CellStyle, int> GenerateCreateWorksheetRowDependencyInfo(
        StringBuilder sb, int typeIndex, EquatableArray<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Any(static x => x.CellStyle is not null));

        sb.AppendLine(FormattableString.Invariant($$"""

                    private static WorksheetRowDependencyInfo CreateWorksheetRowDependencyInfo{{typeIndex}}(Spreadsheet spreadsheet)
                    {
                        var styleIds = new[]
                        {
            """));

        var cellStyleToStyleIdIndex = new Dictionary<CellStyle, int>();

        foreach (var property in properties)
        {
            if (property.CellStyle is not { } style)
                continue;

            if (cellStyleToStyleIdIndex.ContainsKey(style))
                continue;

            cellStyleToStyleIdIndex[style] = cellStyleToStyleIdIndex.Count;

            sb.AppendLine($"""
                            spreadsheet.GetStyleId({style.StyleNameRawString}),
            """);
        }

        sb.AppendLine("""
                        };
                        return new WorksheetRowDependencyInfo { StyleIds = styleIds };
                    }
            """);

        return cellStyleToStyleIdIndex;
    }

    private static void GenerateCellValueMappersArray(StringBuilder sb, RowType rowType)
    {
        if (rowType.Properties.Any(property => property.PropertyCellValueMapper.HasValue))
        {
             sb.AppendLine(FormattableString.Invariant($$"""
                            private static ICellValueMapper[]? _cellValueMappers = new ICellValueMapper[]
                            {
                    """));

            foreach (var cellMapper in rowType.Properties.Where(property => property.PropertyCellValueMapper.HasValue))
            {
                sb.AppendLine(FormattableString.Invariant($$"""
                                new {{cellMapper.PropertyCellValueMapper.GetValueOrDefault().CellValueMapperTypeName}}(),
                   """));
            }

            sb.AppendLine(FormattableString.Invariant($$"""
                            };                     
                    """));
        }
        else
        {
            sb.AppendLine("private static ICellValueMapper[]? _cellValueMappers;");
        }
    }
    
    private static void GenerateAddHeaderRow(StringBuilder sb, int typeIndex, EquatableArray<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine(FormattableString.Invariant($$"""

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
            sb.AppendLine($"""
                        if (obj is null)
                            return spreadsheet.AddRowAsync(ReadOnlyMemory<{rowType.CellType}>.Empty, token);
            """);
        }

        sb.AppendLine("""
                        return AddAsRowInternalAsync(spreadsheet, obj, token);
                    }
            """);
    }

    private static void GenerateArrayPoolRentPart(StringBuilder sb, RowType rowType)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.AppendLine($$"""
                        var cells = ArrayPool<{{rowType.CellType}}>.Shared.Rent({{properties.Count}});
                        try
                        {
            """);
    }

    private static void GenerateGetStyleIdsPart(StringBuilder sb, RowType rowType)
    {
        if (rowType.PropertiesWithStyleAttributes > 0)
        {
            sb.AppendLine($"""
                            var worksheetRowDependencyInfo = spreadsheet.GetOrCreateWorksheetRowDependencyInfo(Default.{rowType.Name});
                            var styleIds = worksheetRowDependencyInfo.StyleIds;
            """);
        }
        else
        {
            sb.AppendLine("""
                            var styleIds = Array.Empty<StyleId>();
            """);
        }
    }

    private static void GenerateArrayPoolReturnPart(StringBuilder sb, RowType rowType)
    {
        sb.AppendLine($$"""
                        }
                        finally
                        {
                            ArrayPool<{{rowType.CellType}}>.Shared.Return(cells, true);
                        }
            """);
    }

    private static void GenerateAddAsRowInternal(StringBuilder sb, RowType rowType)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.AppendLine($$"""

                    private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
                        {{rowType.FullName}} obj,
                        CancellationToken token)
                    {
            """);

        GenerateArrayPoolRentPart(sb, rowType);
        GenerateGetStyleIdsPart(sb, rowType);

        sb.AppendLine("""
                            await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
            """);

        GenerateArrayPoolReturnPart(sb, rowType);

        sb.AppendLine("""
                    }
            """);
    }

    private static void GenerateAddRangeAsRows(StringBuilder sb, RowType rowType)
    {
        sb.AppendLine($$"""

                private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet,
                    IEnumerable<{{rowType.FullNameWithNullableAnnotation}}> objs,
                    CancellationToken token)
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

        sb.AppendLine($$"""

                    private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
                        IEnumerable<{{rowType.FullNameWithNullableAnnotation}}> objs,
                        CancellationToken token)
                    {
            """);

        GenerateArrayPoolRentPart(sb, rowType);
        GenerateGetStyleIdsPart(sb, rowType);

        sb.AppendLine("""
                            foreach (var obj in objs)
                            {
                                await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
                            }
            """);

        GenerateArrayPoolReturnPart(sb, rowType);

        sb.AppendLine("""
                    }
            """);
    }   

    private static void GenerateAddCellsAsRow(StringBuilder sb, RowType rowType, Dictionary<CellStyle, int> cellStyleToStyleIdIndex)
    {
        var properties = rowType.Properties;
        Debug.Assert(properties.Count > 0);

        sb.AppendLine($$"""

                    private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet,
                        {{rowType.FullNameWithNullableAnnotation}} obj,
                        {{rowType.CellType}}[] cells, IReadOnlyList<StyleId> styleIds, CancellationToken token)
                    {
            """);
        
        if (rowType.IsReferenceType)
        {
            sb.AppendLine($"""
                        if (obj is null)
                            return spreadsheet.AddRowAsync(ReadOnlyMemory<{rowType.CellType}>.Empty, token);

            """);
        }

        foreach (var (i, property) in properties.Index())
        {
            sb.AppendLine(FormattableString.Invariant($"""
                        cells[{i}] = {ConstructCell(property, $"obj.{property.Name}", i)};
            """));
        }

        sb.AppendLine($$"""
                        return spreadsheet.AddRowAsync(cells.AsMemory(0, {{properties.Count}}), token);
                    }
            """);

        string ConstructCell(RowTypeProperty property, string value, int propertyIndex)
        {
            int? styleIdIndex = property.CellStyle is { } cellStyle
                ? cellStyleToStyleIdIndex[cellStyle]
                : null;

            var styledCell = rowType.PropertiesWithStyleAttributes > 0;

            if (property.PropertyCellValueMapper.HasValue)
            {
                return FormattableString.Invariant(
                    $"(_cellValueMappers[{propertyIndex}] as ICellValueMapper<{property.PropertyCellValueMapper.GetValueOrDefault().GenericName}>)!.MapToCell({value})");
            }

            return (property.CellValueTruncate, styledCell, styleIdIndex) switch
            {
                ({ } truncate, true, { } i) => FormattableString.Invariant($"ConstructTruncatedStyledCell({value}, {truncate.Value}, styleIds[{i}])"),
                ({ } truncate, true, null) => FormattableString.Invariant($"ConstructTruncatedStyledCell({value}, {truncate.Value}, null)"),
                ({ } truncate, false, _) => FormattableString.Invariant($"ConstructTruncatedDataCell({value}, {truncate.Value})"),
                (null, true, { } i) => FormattableString.Invariant($"new StyledCell({value}, styleIds[{i}])"),
                (null, true, null) => FormattableString.Invariant($"new StyledCell({value}, null)"),
                (null, false, _) => FormattableString.Invariant($"new DataCell({value})"),
            };
        }
    }

    private static void GenerateConstructTruncatedCells(StringBuilder sb)
    {
        sb.AppendLine("""

                    private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
                    {
                        return value is null || value.Length <= truncateLength
                            ? new DataCell(value)
                            : new DataCell(value.AsMemory(0, truncateLength));
                    }

                    private static StyledCell ConstructTruncatedStyledCell(string? value, int truncateLength, StyleId? styleId)
                    {
                        return value is null || value.Length <= truncateLength
                            ? new StyledCell(value, styleId)
                            : new StyledCell(value.AsMemory(0, truncateLength), styleId);
                    }
            """);
    }
}
