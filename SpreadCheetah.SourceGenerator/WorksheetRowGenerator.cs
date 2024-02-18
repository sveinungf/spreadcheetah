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

        foreach (var member in classType.GetMembers())
        {
            if (!member.IsPropertyWithPublicGetter(out var p))
                continue;

            if (!p.Type.IsSupportedType())
            {
                unsupportedPropertyTypeNames.Add(p.Type.Name);
                continue;
            }

            TypedConstant? columnHeaderAttributeValue = null;
            ColumnOrder? columnOrder = null;

            foreach (var attribute in p.GetAttributes())
            {
                if (columnHeaderAttributeValue is null && attribute.TryParseColumnHeaderAttribute(out var arg))
                    columnHeaderAttributeValue = arg;

                if (columnOrder is null && attribute.TryParseColumnOrderAttribute(token, out var orderArg))
                    columnOrder = orderArg;
            }

            var rowTypeProperty = p.ToRowTypeProperty(columnHeaderAttributeValue);

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
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("using SpreadCheetah;");
        sb.AppendLine("using SpreadCheetah.SourceGeneration;");
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Buffers;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using System.Threading.Tasks;");
        sb.AppendLine();
    }

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass, SourceProductionContext context)
    {
        GenerateHeader(sb);

        if (contextClass.Namespace is { } ns)
            sb.AppendLine($"namespace {ns}");

        var accessibility = SyntaxFacts.GetText(contextClass.DeclaredAccessibility);

        sb.AppendLine("{");
        sb.AppendLine($"    {accessibility} partial class {contextClass.Name}");
        sb.AppendLine("    {");
        sb.AppendLine($"        private static {contextClass.Name}? _default;");
        sb.AppendLine($"        public static {contextClass.Name} Default => _default ??= new {contextClass.Name}();");
        sb.AppendLine();
        sb.AppendLine($"        public {contextClass.Name}()");
        sb.AppendLine("        {");
        sb.AppendLine("        }");

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

        sb.AppendLine("    }");
        sb.AppendLine("}");
    }

    private static void GenerateCodeForType(StringBuilder sb, int typeIndex, RowType rowType,
        ContextClass contextClass, SourceProductionContext context)
    {
        ReportDiagnostics(rowType, rowType.WorksheetRowAttributeLocation, contextClass.Options, context);

        sb.AppendLine().AppendLine(FormattableString.Invariant($$"""
                private WorksheetRowTypeInfo<{{rowType.FullName}}>? _{{rowType.Name}};
                public WorksheetRowTypeInfo<{{rowType.FullName}}> {{rowType.Name}} => _{{rowType.Name}}
        """));

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
        GenerateAddAsRow(sb, 2, rowType);
        GenerateAddRangeAsRows(sb, 2, rowType);
        GenerateAddAsRowInternal(sb, 2, rowType.FullName, rowType.Properties);
        GenerateAddRangeAsRowsInternal(sb, rowType, rowType.Properties);
        GenerateAddEnumerableAsRows(sb, 2, rowType);
        GenerateAddCellsAsRow(sb, 2, rowType, rowType.Properties);
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

    private static void GenerateAddHeaderRow(StringBuilder sb, int typeIndex, IReadOnlyCollection<RowTypeProperty> properties)
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
            sb.AppendLine(FormattableString.Invariant($"""
                            cells[{i}] = new StyledCell({property.ColumnHeader}, styleId);
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

    private static void GenerateAddAsRow(StringBuilder sb, int indent, RowType rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, ")
            .Append(rowType.FullNameWithNullableAnnotation)
            .AppendLine(" obj, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    if (spreadsheet is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(spreadsheet));");

        if (rowType.IsReferenceType)
        {
            sb.AppendLine(indent + 1, "if (obj is null)");
            sb.AppendLine(indent + 1, "    return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
        }

        sb.AppendLine(indent, "    return AddAsRowInternalAsync(spreadsheet, obj, token);");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddAsRowInternal(StringBuilder sb, int indent, string rowTypeFullname, IReadOnlyCollection<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine();
        sb.AppendLine(indent, $"private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, {rowTypeFullname} obj, CancellationToken token)");
        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, $"    var cells = ArrayPool<DataCell>.Shared.Rent({properties.Count});");
        sb.AppendLine(indent, "    try");
        sb.AppendLine(indent, "    {");
        sb.AppendLine(indent, "        await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "    finally");
        sb.AppendLine(indent, "    {");
        sb.AppendLine(indent, "        ArrayPool<DataCell>.Shared.Return(cells, true);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddRangeAsRows(StringBuilder sb, int indent, RowType rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<")
            .Append(rowType.FullNameWithNullableAnnotation)
            .AppendLine("> objs, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    if (spreadsheet is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(spreadsheet));");
        sb.AppendLine(indent, "    if (objs is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(objs));");
        sb.AppendLine(indent, "    return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddRangeAsRowsInternal(StringBuilder sb, RowType rowType, IReadOnlyCollection<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        var typeString = rowType.FullNameWithNullableAnnotation;
        sb.Append($$"""

                private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<{{typeString}}> objs, CancellationToken token)
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

    private static void GenerateAddEnumerableAsRows(StringBuilder sb, int indent, RowType rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static async ValueTask AddEnumerableAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<")
            .Append(rowType.FullNameWithNullableAnnotation)
            .AppendLine("> objs, DataCell[] cells, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    foreach (var obj in objs)");
        sb.AppendLine(indent, "    {");
        sb.AppendLine(indent, "        await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddCellsAsRow(StringBuilder sb, int indent, RowType rowType, IReadOnlyCollection<RowTypeProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, ")
            .Append(rowType.FullNameWithNullableAnnotation)
            .AppendLine(" obj, DataCell[] cells, CancellationToken token)");

        sb.AppendLine(indent, "{");

        if (rowType.IsReferenceType)
        {
            sb.AppendLine(indent, "    if (obj is null)");
            sb.AppendLine(indent, "        return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
            sb.AppendLine();
        }

        foreach (var (i, property) in properties.Index())
        {
            sb.AppendIndentation(indent + 1)
                .Append("cells[")
                .Append(i)
                .Append("] = new DataCell(obj.")
                .Append(property.Name)
                .AppendLine(");");
        }

        sb.AppendLine(indent, $"    return spreadsheet.AddRowAsync(cells.AsMemory(0, {properties.Count}), token);");
        sb.AppendLine(indent, "}");
    }
}
