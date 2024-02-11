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

        var rowTypes = new Dictionary<INamedTypeSymbol, Location>(SymbolEqualityComparer.Default);
        GeneratorOptions? generatorOptions = null;

        foreach (var worksheetRowAttribute in context.Attributes)
        {
            if (TryParseWorksheetRowAttribute(worksheetRowAttribute, token, out var typeSymbol, out var location)
                && !rowTypes.ContainsKey(typeSymbol))
            {
                rowTypes[typeSymbol] = location;
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
        var unsupportedPropertyNames = new List<IPropertySymbol>();

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

            if (!IsSupportedType(p.Type, supportedNullableTypes))
            {
                unsupportedPropertyNames.Add(p);
                continue;
            }

            var columnHeader = GetColumnHeader(p);
            var columnProperty = new ColumnProperty(p.Name, columnHeader);

            if (!TryGetExplicitColumnOrder(p, context.CancellationToken, out var columnOrder, out var location))
                implicitOrderProperties.Add(columnProperty);
            else if (!explicitOrderProperties.ContainsKey(columnOrder))
                explicitOrderProperties.Add(columnOrder, columnProperty);
            else
                context.ReportDiagnostic(Diagnostic.Create(Diagnostics.DuplicateColumnOrder, location, classType.Name));
        }

        explicitOrderProperties.AddWithImplicitKeys(implicitOrderProperties);

        return new TypePropertiesInfo(explicitOrderProperties, unsupportedPropertyNames);
    }

    private static string GetColumnHeader(IPropertySymbol property)
    {
        foreach (var attribute in property.GetAttributes())
        {
            if (TryParseColumnHeaderAttribute(attribute, out var arg))
                return arg.ToCSharpString();
        }

        return @$"""{property.Name}""";
    }

    private static bool TryGetExplicitColumnOrder(IPropertySymbol property,
        CancellationToken token, out int columnOrder, out Location? location)
    {
        columnOrder = 0;
        location = null;

        foreach (var attribute in property.GetAttributes())
        {
            if (!TryParseColumnOrderAttribute(attribute, out columnOrder))
                continue;

            location = attribute.ApplicationSyntaxReference?.GetSyntax(token).GetLocation();
            return true;
        }

        return false;
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

    private static bool IsSupportedType(ITypeSymbol type, EquatableArray<string> supportedNullableTypes)
    {
        return type.SpecialType == SpecialType.System_String
            || SupportedPrimitiveTypes.Contains(type.SpecialType)
            || IsSupportedNullableType(type, supportedNullableTypes);
    }

    private static bool IsSupportedNullableType(ITypeSymbol type, EquatableArray<string> supportedNullableTypes)
    {
        return type.NullableAnnotation == NullableAnnotation.Annotated
            && supportedNullableTypes.Contains(type.ToDisplayString(), StringComparer.Ordinal);
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

    private static void GenerateCodeForType(StringBuilder sb, int typeIndex, INamedTypeSymbol rowType, Location location,
        ContextClass contextClass, EquatableArray<string> supportedNullableTypes, SourceProductionContext context)
    {
        var rowTypeName = rowType.Name;
        var rowTypeFullName = rowType.ToString();

        var info = AnalyzeTypeProperties(supportedNullableTypes, rowType, context);
        ReportDiagnostics(info, rowType, location, contextClass.Options, context);

        sb.AppendLine().AppendLine(FormattableString.Invariant($$"""
                private WorksheetRowTypeInfo<{{rowTypeFullName}}>? _{{rowTypeName}};
                public WorksheetRowTypeInfo<{{rowTypeFullName}}> {{rowTypeName}} => _{{rowTypeName}}
        """));

        if (info.Properties.Count == 0)
        {
            sb.AppendLine($$"""
                        ??= EmptyWorksheetRowContext.CreateTypeInfo<{{rowTypeFullName}}>();
            """);

            return;
        }

        sb.AppendLine(FormattableString.Invariant($$"""
                        ??= WorksheetRowMetadataServices.CreateObjectInfo<{{rowTypeFullName}}>(AddHeaderRow{{typeIndex}}Async, AddAsRowAsync, AddRangeAsRowsAsync);
            """));

        var properties = info.Properties.Values.ToList();
        GenerateAddHeaderRow(sb, typeIndex, properties);
        GenerateAddAsRow(sb, 2, rowType);
        GenerateAddRangeAsRows(sb, 2, rowType);
        GenerateAddAsRowInternal(sb, 2, rowTypeFullName, properties);
        GenerateAddRangeAsRowsInternal(sb, rowType, properties);
        GenerateAddEnumerableAsRows(sb, 2, rowType);
        GenerateAddCellsAsRow(sb, 2, rowType, properties);
    }

    private static void ReportDiagnostics(TypePropertiesInfo info, INamedTypeSymbol rowType, Location location, GeneratorOptions? options, SourceProductionContext context)
    {
        if (options?.SuppressWarnings ?? false) return;

        if (info.Properties.Count == 0)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.NoPropertiesFound, location, rowType.Name));

        if (info.UnsupportedProperties.FirstOrDefault() is { } unsupportedProperty)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.UnsupportedTypeForCellValue, location, rowType.Name, unsupportedProperty.Type.Name));
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

    private static void GenerateAddAsRow(StringBuilder sb, int indent, INamedTypeSymbol rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, ")
            .AppendType(rowType)
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

    private static void GenerateAddRangeAsRows(StringBuilder sb, int indent, INamedTypeSymbol rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<")
            .AppendType(rowType)
            .AppendLine("> objs, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    if (spreadsheet is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(spreadsheet));");
        sb.AppendLine(indent, "    if (objs is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(objs));");
        sb.AppendLine(indent, "    return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddRangeAsRowsInternal(StringBuilder sb, INamedTypeSymbol rowType, IReadOnlyCollection<ColumnProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        var typeString = rowType.ToTypeString();
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

    private static void GenerateAddEnumerableAsRows(StringBuilder sb, int indent, INamedTypeSymbol rowType)
    {
        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static async ValueTask AddEnumerableAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<")
            .AppendType(rowType)
            .AppendLine("> objs, DataCell[] cells, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    foreach (var obj in objs)");
        sb.AppendLine(indent, "    {");
        sb.AppendLine(indent, "        await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddCellsAsRow(StringBuilder sb, int indent, INamedTypeSymbol rowType, IReadOnlyList<ColumnProperty> properties)
    {
        Debug.Assert(properties.Count > 0);

        sb.AppendLine()
            .AppendIndentation(indent)
            .Append("private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, ")
            .AppendType(rowType)
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
