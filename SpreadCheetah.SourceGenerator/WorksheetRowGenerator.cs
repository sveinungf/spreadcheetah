using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace SpreadCheetah.SourceGenerators;

[Generator]
public class WorksheetRowGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // TODO: Add tracking name (if it can't be hardcoded)
        var supportedNullableTypes = context.CompilationProvider.Select(static (c, _) => GetSupportedNullableTypes(c));

        var contextClasses = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                "SpreadCheetah.SourceGeneration.WorksheetRowAttribute",
                IsSyntaxTargetForGeneration,
                GetSemanticTargetForGeneration)
            .Where(static x => x is not null)
            .WithTrackingName(TrackingNames.InitialExtraction);

        var combined = contextClasses
            .Combine(supportedNullableTypes)
            .WithTrackingName(TrackingNames.Transform);

        context.RegisterSourceOutput(combined, static (spc, source) => Execute(source.Left, source.Right, spc));
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

        var rowTypes = new Dictionary<INamedTypeSymbol, LocationInfo>(SymbolEqualityComparer.Default);
        GeneratorOptions? generatorOptions = null;

        foreach (var worksheetRowAttribute in context.Attributes)
        {
            if (TryParseWorksheetRowAttribute(worksheetRowAttribute, token, out var typeSymbol, out var location)
                && !rowTypes.ContainsKey(typeSymbol))
            {
                var locationInfo = location.ToLocationInfo();
                if (locationInfo is not null)
                    rowTypes[typeSymbol] = locationInfo;
            }
        }

        if (rowTypes.Count == 0)
            return null;

        foreach (var attribute in classSymbol.GetAttributes())
        {
            if (TryParseOptionsAttribute(attribute, out var options))
                generatorOptions = options;
        }

        return new ContextClass(
            DeclaredAccessibility: classSymbol.DeclaredAccessibility,
            Namespace: classSymbol.ContainingNamespace is { IsGlobalNamespace: false } ns ? ns.ToString() : null,
            Name: classSymbol.Name,
            RowTypes: rowTypes,
            Options: generatorOptions);
    }

    private static bool TryParseWorksheetRowAttribute(
        AttributeData attribute,
        CancellationToken token,
        [NotNullWhen(true)] out INamedTypeSymbol? typeSymbol,
        [NotNullWhen(true)] out Location? location)
    {
        typeSymbol = null;
        location = null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: INamedTypeSymbol symbol }])
            return false;

        if (symbol.Kind == SymbolKind.ErrorType)
            return false;

        var syntaxReference = attribute.ApplicationSyntaxReference;
        if (syntaxReference is null)
            return false;

        location = syntaxReference.GetSyntax(token).GetLocation();
        typeSymbol = symbol;
        return true;
    }

    private static bool TryParseOptionsAttribute(
        AttributeData attribute,
        [NotNullWhen(true)] out GeneratorOptions? options)
    {
        options = null;

        if (!string.Equals(Attributes.GenerationOptions, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return false;

        if (attribute.NamedArguments.IsDefaultOrEmpty)
            return false;

        foreach (var (key, value) in attribute.NamedArguments)
        {
            if (!string.Equals(key, "SuppressWarnings", StringComparison.Ordinal))
                continue;

            if (value.Value is bool suppressWarnings)
            {
                options = new GeneratorOptions(suppressWarnings);
                return true;
            }
        }

        return false;
    }

    private static bool TryParseColumnHeaderAttribute(
        AttributeData attribute,
        out TypedConstant attributeArg)
    {
        attributeArg = default;

        if (!string.Equals(Attributes.ColumnHeader, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: string } arg])
            return false;

        attributeArg = arg;
        return true;
    }

    private static bool TryParseColumnOrderAttribute(
        AttributeData attribute,
        out int order)
    {
        order = 0;

        if (!string.Equals(Attributes.ColumnOrder, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        order = attributeValue;
        return true;
    }

    private static TypePropertiesInfo AnalyzeTypeProperties(EquatableArray<string> supportedNullableTypes,
        ITypeSymbol classType, SourceProductionContext context)
    {
        var implicitOrderProperties = new List<ColumnProperty>();
        var explicitOrderProperties = new SortedDictionary<int, ColumnProperty>();
        var unsupportedPropertyTypeNames = new HashSet<string>(StringComparer.Ordinal);

        foreach (var member in classType.GetMembers())
        {
            if (member is not IPropertySymbol
                {
                    DeclaredAccessibility: Accessibility.Public,
                    IsStatic: false,
                    IsWriteOnly: false
                } p)
            {
                continue;
            }

            TypedConstant? columnHeaderAttributeValue = null;
            int? columnOrderValue = null;
            Location? columnOrderAttributeLocation = null;

            foreach (var attribute in p.GetAttributes())
            {
                if (columnHeaderAttributeValue is null && TryParseColumnHeaderAttribute(attribute, out var arg))
                {
                    columnHeaderAttributeValue = arg;
                }

                if (columnOrderValue is null && TryParseColumnOrderAttribute(attribute, out var order))
                {
                    columnOrderValue = order;
                    columnOrderAttributeLocation = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation();
                }
            }

            var rowTypeProperty = new RowTypeProperty(
                Name: p.Name,
                TypeName: p.Type.Name,
                TypeFullName: p.Type.ToDisplayString(),
                TypeNullableAnnotation: p.NullableAnnotation,
                TypeSpecialType: p.Type.SpecialType,
                ColumnHeaderAttributeValue: columnHeaderAttributeValue);

            if (!IsSupportedType(rowTypeProperty, supportedNullableTypes))
            {
                unsupportedPropertyTypeNames.Add(rowTypeProperty.TypeName);
                continue;
            }

            var columnHeader = GetColumnHeader(rowTypeProperty);
            var columnProperty = new ColumnProperty(rowTypeProperty.Name, columnHeader);

            if (columnOrderValue is not { } columnOrder)
                implicitOrderProperties.Add(columnProperty);
            else if (!explicitOrderProperties.ContainsKey(columnOrder))
                explicitOrderProperties.Add(columnOrder, columnProperty);
            else
                context.ReportDiagnostic(Diagnostic.Create(Diagnostics.DuplicateColumnOrder, columnOrderAttributeLocation, classType.Name));
        }

        explicitOrderProperties.AddWithImplicitKeys(implicitOrderProperties);

        return new TypePropertiesInfo(explicitOrderProperties, new EquatableArray<string>(unsupportedPropertyTypeNames.ToArray()));
    }

    private static string GetColumnHeader(RowTypeProperty property)
    {
        return property.ColumnHeaderAttributeValue is { } attributeValue
            ? attributeValue.ToCSharpString()
            : @$"""{property.Name}""";
    }

    private static EquatableArray<string> GetSupportedNullableTypes(Compilation compilation)
    {
        var result = new List<string>();
        var nullableT = compilation.GetTypeByMetadataName("System.Nullable`1");

        // TODO: Could be hardcoded?
        foreach (var primitiveType in SupportedPrimitiveTypes)
        {
            var nullableType = nullableT?.Construct(compilation.GetSpecialType(primitiveType));
            if (nullableType is not null)
                result.Add(nullableType.ToDisplayString());
        }

        return new EquatableArray<string>(result.ToArray());
    }

    private static bool IsSupportedType(RowTypeProperty typeProperty, EquatableArray<string> supportedNullableTypes)
    {
        return typeProperty.TypeSpecialType == SpecialType.System_String
            || SupportedPrimitiveTypes.Contains(typeProperty.TypeSpecialType)
            || IsSupportedNullableType(typeProperty, supportedNullableTypes);
    }

    private static bool IsSupportedNullableType(RowTypeProperty typeProperty, EquatableArray<string> supportedNullableTypes)
    {
        return typeProperty.TypeNullableAnnotation == NullableAnnotation.Annotated
            && supportedNullableTypes.Contains(typeProperty.TypeFullName, StringComparer.Ordinal);
    }

    private static readonly SpecialType[] SupportedPrimitiveTypes =
    [
        SpecialType.System_Boolean,
        SpecialType.System_DateTime,
        SpecialType.System_Decimal,
        SpecialType.System_Double,
        SpecialType.System_Int32,
        SpecialType.System_Int64,
        SpecialType.System_Single
    ];

    private static void Execute(ContextClass? contextClass, EquatableArray<string> supportedNullableTypes, SourceProductionContext context)
    {
        if (contextClass is null)
            return;

        var sb = new StringBuilder();
        GenerateCode(sb, contextClass, supportedNullableTypes, context);

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

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass, EquatableArray<string> supportedNullableTypes, SourceProductionContext context)
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
        foreach (var (rowType, location) in contextClass.RowTypes)
        {
            var rowTypeName = rowType.Name;
            if (!rowTypeNames.Add(rowTypeName))
                continue;

            GenerateCodeForType(sb, typeIndex, rowType, location, contextClass, supportedNullableTypes, context);
            ++typeIndex;
        }

        sb.AppendLine("    }");
        sb.AppendLine("}");
    }

    private static void GenerateCodeForType(StringBuilder sb, int typeIndex, INamedTypeSymbol rowTypeOld, LocationInfo location,
        ContextClass contextClass, EquatableArray<string> supportedNullableTypes, SourceProductionContext context)
    {
        var rowType = new RowType(
            Name: rowTypeOld.Name,
            FullName: rowTypeOld.ToString(),
            FullNameWithNullableAnnotation: rowTypeOld.IsReferenceType ? $"{rowTypeOld}?" : rowTypeOld.ToString(),
            IsReferenceType: rowTypeOld.IsReferenceType);

        var info = AnalyzeTypeProperties(supportedNullableTypes, rowTypeOld, context);
        ReportDiagnostics(info, rowType, location, contextClass.Options, context);

        sb.AppendLine().AppendLine(FormattableString.Invariant($$"""
                private WorksheetRowTypeInfo<{{rowType.FullName}}>? _{{rowType.Name}};
                public WorksheetRowTypeInfo<{{rowType.FullName}}> {{rowType.Name}} => _{{rowType.Name}}
        """));

        if (info.Properties.Count == 0)
        {
            sb.AppendLine($$"""
                        ??= EmptyWorksheetRowContext.CreateTypeInfo<{{rowType.FullName}}>();
            """);

            return;
        }

        sb.AppendLine(FormattableString.Invariant($$"""
                        ??= WorksheetRowMetadataServices.CreateObjectInfo<{{rowType.FullName}}>(AddHeaderRow{{typeIndex}}Async, AddAsRowAsync, AddRangeAsRowsAsync);
            """));

        var properties = info.Properties.Values.ToList();
        GenerateAddHeaderRow(sb, typeIndex, properties);
        GenerateAddAsRow(sb, 2, rowType);
        GenerateAddRangeAsRows(sb, 2, rowType);
        GenerateAddAsRowInternal(sb, 2, rowType.FullName, properties);
        GenerateAddRangeAsRowsInternal(sb, rowType, properties);
        GenerateAddEnumerableAsRows(sb, 2, rowType);
        GenerateAddCellsAsRow(sb, 2, rowType, properties);
    }

    private static void ReportDiagnostics(TypePropertiesInfo info, RowType rowType, LocationInfo location, GeneratorOptions? options, SourceProductionContext context)
    {
        if (options?.SuppressWarnings ?? false) return;

        if (info.Properties.Count == 0)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.NoPropertiesFound, location.ToLocation(), rowType.Name));

        if (info.UnsupportedPropertyTypeNames.FirstOrDefault() is { } unsupportedPropertyTypeName)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.UnsupportedTypeForCellValue, location.ToLocation(), rowType.Name, unsupportedPropertyTypeName));
    }

    private static void GenerateAddHeaderRow(StringBuilder sb, int typeIndex, IReadOnlyList<ColumnProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine().AppendLine(FormattableString.Invariant($$"""
                private static async ValueTask AddHeaderRow{{typeIndex}}Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
                {
                    var cells = ArrayPool<StyledCell>.Shared.Rent({{properties.Count}});
                    try
                    {
        """));

        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
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

    private static void GenerateAddAsRowInternal(StringBuilder sb, int indent, string rowTypeFullname, IReadOnlyCollection<ColumnProperty> properties)
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

    private static void GenerateAddRangeAsRowsInternal(StringBuilder sb, RowType rowType, IReadOnlyCollection<ColumnProperty> properties)
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

    private static void GenerateAddCellsAsRow(StringBuilder sb, int indent, RowType rowType, IReadOnlyList<ColumnProperty> properties)
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

        for (var i = 0; i < properties.Count; i++)
        {
            sb.AppendIndentation(indent + 1)
                .Append("cells[")
                .Append(i)
                .Append("] = new DataCell(obj.")
                .Append(properties[i].PropertyName)
                .AppendLine(");");
        }

        sb.AppendLine(indent, $"    return spreadsheet.AddRowAsync(cells.AsMemory(0, {properties.Count}), token);");
        sb.AppendLine(indent, "}");
    }
}
