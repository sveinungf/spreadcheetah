using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
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
            if (!attribute.TryParseWorksheetRowAttribute(out var typeSymbol))
                continue;

            var rowType = AnalyzeTypeProperties(typeSymbol, token);
            if (!rowTypes.Exists(x => string.Equals(x.FullName, rowType.FullName, StringComparison.Ordinal)))
                rowTypes.Add(rowType);
        }

        if (rowTypes.Count == 0)
            return null;

        return new ContextClass(
            DeclaredAccessibility: classSymbol.DeclaredAccessibility,
            Namespace: classSymbol.ContainingNamespace is { IsGlobalNamespace: false } ns ? ns.ToString() : null,
            Name: classSymbol.Name,
            RowTypes: rowTypes.ToEquatableArray());
    }

    private static RowType AnalyzeTypeProperties(ITypeSymbol rowType, CancellationToken token)
    {
        var implicitOrderProperties = new List<RowTypeProperty>();
        var explicitOrderProperties = new SortedDictionary<int, RowTypeProperty>();
        var hasStyleAttributes = false;
        var analyzer = new PropertyAnalyzer(NullDiagnosticsReporter.Instance);

        var properties = rowType
            .GetClassAndBaseClassProperties()
            .Where(x => x.IsInstancePropertyWithPublicGetter());

        foreach (var property in properties)
        {
            var data = analyzer.Analyze(property, token);

            if (data.CellValueConverter is null && !property.Type.IsSupportedType())
                continue;

            var rowTypeProperty = new RowTypeProperty(
                Name: property.Name,
                CellFormat: data.CellFormat,
                CellStyle: data.CellStyle,
                CellValueConverter: data.CellValueConverter,
                CellValueTruncate: data.CellValueTruncate,
                ColumnHeader: data.ColumnHeader?.ToColumnHeaderInfo(),
                ColumnWidth: data.ColumnWidth);

            if (rowTypeProperty.HasStyle)
                hasStyleAttributes = true;

            if (data.ColumnOrder is not { } order)
                implicitOrderProperties.Add(rowTypeProperty);
            else if (!explicitOrderProperties.ContainsKey(order.Value))
                explicitOrderProperties.Add(order.Value, rowTypeProperty);
        }

        explicitOrderProperties.AddWithImplicitKeys(implicitOrderProperties);

        return new RowType(
            FullName: rowType.ToString(),
            HasStyleAttributes: hasStyleAttributes,
            IsReferenceType: rowType.IsReferenceType,
            Name: rowType.Name,
            Properties: explicitOrderProperties.Values.ToEquatableArray());
    }

    private static void Execute(ContextClass? contextClass, SourceProductionContext context)
    {
        if (contextClass is null)
            return;

        var sb = new StringBuilder();
        GenerateCode(sb, contextClass);

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

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass)
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

        var cellValueConverters = GenerateCellValueConverters(sb, contextClass.RowTypes);

        var rowTypeNames = new HashSet<string>(StringComparer.Ordinal);

        var typeIndex = 0;
        foreach (var rowType in contextClass.RowTypes)
        {
            var rowTypeName = rowType.Name;
            if (!rowTypeNames.Add(rowTypeName))
                continue;

            GenerateCodeForType(sb, typeIndex, cellValueConverters, rowType);
            ++typeIndex;
        }

        GenerateConstructTruncatedDataCell(sb);

        sb.AppendLine("""
                }
            }
            """);
    }

    private static void GenerateCodeForType(StringBuilder sb, int typeIndex,
        Dictionary<string, string> cellValueConverters, RowType rowType)
    {
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
        var doGenerateCreateWorksheetRowDependencyInfo = properties.Any(static x => x.HasStyle);

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

        var styleLookup = doGenerateCreateWorksheetRowDependencyInfo
            ? GenerateCreateWorksheetRowDependencyInfo(sb, typeIndex, properties)
            : null;

        GenerateAddHeaderRow(sb, typeIndex, properties);
        GenerateAddAsRow(sb, rowType);
        GenerateAddRangeAsRows(sb, rowType);
        GenerateAddAsRowInternal(sb, rowType);
        GenerateAddRangeAsRowsInternal(sb, rowType);
        GenerateAddCellsAsRow(sb, rowType, styleLookup, cellValueConverters);
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

    private static StyleLookup GenerateCreateWorksheetRowDependencyInfo(
        StringBuilder sb, int typeIndex, EquatableArray<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Any(static x => x.HasStyle));

        sb.AppendLine(FormattableString.Invariant($$"""

                    private static WorksheetRowDependencyInfo CreateWorksheetRowDependencyInfo{{typeIndex}}(Spreadsheet spreadsheet)
                    {
                        var styleIds = new[]
                        {
            """));

        var lookup = new StyleLookup();

        foreach (var property in properties)
        {
            _ = property switch
            {
                { CellFormat: { } format } => HandleCellFormat(sb, lookup, format),
                { CellStyle: { } style } => HandleCellStyle(sb, lookup, style),
                _ => false
            };
        }

        sb.AppendLine("""
                        };
                        return new WorksheetRowDependencyInfo(styleIds);
                    }
            """);

        return lookup;
    }

    private static bool HandleCellFormat(StringBuilder sb, StyleLookup styleLookup, CellFormat format)
    {
        if (!styleLookup.TryAdd(format))
            return false;

        if (format.RawString is { } rawString)
        {
            sb.AppendLine($$"""
                            spreadsheet.AddStyle(new Style { Format = NumberFormat.Custom({{rawString}}) }),
            """);
        }
        else if (format.StandardFormat is { } standardFormat)
        {
            sb.AppendLine($$"""
                            spreadsheet.AddStyle(new Style { Format = NumberFormat.Standard(StandardNumberFormat.{{standardFormat}}) }),
            """);
        }

        return true;
    }

    private static bool HandleCellStyle(StringBuilder sb, StyleLookup styleLookup, CellStyle style)
    {
        if (!styleLookup.TryAdd(style))
            return false;

        sb.AppendLine($"""
                            spreadsheet.GetStyleId({style.StyleNameRawString}),
            """);
        return true;
    }

    private static Dictionary<string, string> GenerateCellValueConverters(StringBuilder sb, EquatableArray<RowType> rowTypes)
    {
        var converterTypeToFieldMap = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var rowType in rowTypes)
        {
            foreach (var property in rowType.Properties)
            {
                if (property.CellValueConverter is not { } cellValueConverter)
                    continue;

                var converterTypeName = cellValueConverter.ConverterTypeName;
                if (converterTypeToFieldMap.ContainsKey(converterTypeName))
                    continue;

                var fieldName = FormattableString.Invariant($"_valueConverter{converterTypeToFieldMap.Count}");
                converterTypeToFieldMap[converterTypeName] = fieldName;

                sb.AppendLine($$"""
                           private static readonly {{converterTypeName}} {{fieldName}} = new {{converterTypeName}}(); 
                   """);
            }
        }

        return converterTypeToFieldMap;
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
        if (rowType.HasStyleAttributes)
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

    private static void GenerateAddCellsAsRow(StringBuilder sb, RowType rowType,
        StyleLookup? styleLookup, Dictionary<string, string> valueConverters)
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
            int? styleIdIndex = null;
            _ = property switch
            {
                _ when styleLookup is null => false,
                { CellFormat: { } format } => styleLookup.TryGetStyleIdIndex(format, out styleIdIndex),
                { CellStyle: { } style } => styleLookup.TryGetStyleIdIndex(style, out styleIdIndex),
                _ => false
            };

            sb.AppendLine(FormattableString.Invariant($"""
                        cells[{i}] = {ConstructCell(property, styleIdIndex)};
            """));
        }

        sb.AppendLine($$"""
                        return spreadsheet.AddRowAsync(cells.AsMemory(0, {{properties.Count}}), token);
                    }
            """);

        string ConstructCell(RowTypeProperty property, int? styleIdIndex)
        {
            var value = $"obj.{property.Name}";

            var constructDataCell = (property.CellValueConverter, property.CellValueTruncate) switch
            {
                (null, null) => $"new DataCell({value})",
                ({ } converter, _) => $"{valueConverters[converter.ConverterTypeName]}.ConvertToDataCell({value})",
                (null, { } truncate) => FormattableString.Invariant($"ConstructTruncatedDataCell({value}, {truncate.Value})")
            };

            var styleId = (rowType.HasStyleAttributes, styleIdIndex) switch
            {
                (true, { } i) => FormattableString.Invariant($"styleIds[{i}]"),
                (true, _) => "null",
                _ => null
            };

            return styleId is null
                ? constructDataCell
                : $"new StyledCell({constructDataCell}, {styleId})";
        }
    }

    private static void GenerateConstructTruncatedDataCell(StringBuilder sb)
    {
        sb.AppendLine("""

                    private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
                    {
                        return value is null || value.Length <= truncateLength
                            ? new DataCell(value)
                            : new DataCell(value.AsMemory(0, truncateLength));
                    }
            """);
    }
}
