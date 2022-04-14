using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SpreadCheetah.SourceGenerator;
using SpreadCheetah.SourceGenerator.Helpers;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace SpreadCheetah.SourceGenerators;

[Generator]
public class WorksheetRowGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var filtered = context.SyntaxProvider
            .CreateSyntaxProvider(
                static (s, _) => IsSyntaxTargetForGeneration(s),
                static (ctx, token) => GetSemanticTargetForGeneration(ctx, token))
            .Where(static x => x is not null)
            .Collect();

        var source = context.CompilationProvider.Combine(filtered);

        context.RegisterSourceOutput(source, static (spc, source) => Execute(source.Left, source.Right, spc));
    }

    private static bool IsSyntaxTargetForGeneration(SyntaxNode syntaxNode) => syntaxNode is ClassDeclarationSyntax
    {
        AttributeLists.Count: > 0,
        BaseList.Types.Count: > 0
    };

    private static INamedTypeSymbol? GetWorksheetRowAttributeType(Compilation compilation)
    {
        return compilation.GetTypeByMetadataName("SpreadCheetah.SourceGeneration.WorksheetRowAttribute");
    }

    private static INamedTypeSymbol? GetGenerationOptionsAttributeType(Compilation compilation)
    {
        return compilation.GetTypeByMetadataName("SpreadCheetah.SourceGeneration.WorksheetRowGenerationOptionsAttribute");
    }

    private static INamedTypeSymbol? GetContextBaseType(Compilation compilation)
    {
        return compilation.GetTypeByMetadataName("SpreadCheetah.SourceGeneration.WorksheetRowContext");
    }

    private static ContextClass? GetSemanticTargetForGeneration(GeneratorSyntaxContext context, CancellationToken token)
    {
        if (context.Node is not ClassDeclarationSyntax classDeclaration)
            return null;

        if (!classDeclaration.Modifiers.Any(x => x.IsKind(SyntaxKind.PartialKeyword)))
            return null;

        var classSymbol = context.SemanticModel.GetDeclaredSymbol(classDeclaration, token);
        if (classSymbol is null)
            return null;

        if (classSymbol.IsStatic)
            return null;

        var baseType = classSymbol.BaseType;
        if (baseType is null)
            return null;

        var baseContext = GetContextBaseType(context.SemanticModel.Compilation);
        if (baseContext is null)
            return null;

        if (!SymbolEqualityComparer.Default.Equals(baseContext, baseType))
            return null;

        var worksheetRowAttribute = GetWorksheetRowAttributeType(context.SemanticModel.Compilation);
        if (worksheetRowAttribute is null)
            return null;

        var optionsAttribute = GetGenerationOptionsAttributeType(context.SemanticModel.Compilation);
        if (optionsAttribute is null)
            return null;

        var rowTypes = new Dictionary<INamedTypeSymbol, Location>(SymbolEqualityComparer.Default);
        GeneratorOptions? generatorOptions = null;

        foreach (var attribute in classSymbol.GetAttributes())
        {
            if (TryParseWorksheetRowAttribute(attribute, worksheetRowAttribute, token, out var typeSymbol, out var location)
                && !rowTypes.ContainsKey(typeSymbol))
            {
                rowTypes[typeSymbol] = location;
                continue;
            }

            if (TryParseOptionsAttribute(attribute, optionsAttribute, out var options))
                generatorOptions = options;
        }

        return rowTypes.Count > 0
            ? new ContextClass(classSymbol, rowTypes, generatorOptions)
            : null;
    }

    private static bool TryParseWorksheetRowAttribute(
        AttributeData attribute,
        INamedTypeSymbol expectedAttribute,
        CancellationToken token,
        [NotNullWhen(true)] out INamedTypeSymbol? typeSymbol,
        [NotNullWhen(true)] out Location? location)
    {
        typeSymbol = null;
        location = null;

        if (!SymbolEqualityComparer.Default.Equals(expectedAttribute, attribute.AttributeClass))
            return false;

        var args = attribute.ConstructorArguments;
        if (args.Length != 1)
            return false;

        if (args[0].Value is not INamedTypeSymbol symbol)
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
        INamedTypeSymbol expectedAttribute,
        [NotNullWhen(true)] out GeneratorOptions? options)
    {
        options = null;

        if (!SymbolEqualityComparer.Default.Equals(expectedAttribute, attribute.AttributeClass))
            return false;

        if (attribute.NamedArguments.IsDefaultOrEmpty)
            return false;

        foreach (var arg in attribute.NamedArguments)
        {
            if (!string.Equals(arg.Key, "SuppressWarnings", StringComparison.Ordinal))
                continue;

            if (arg.Value.Value is bool suppressWarnings)
            {
                options = new GeneratorOptions(suppressWarnings);
                return true;
            }
        }

        return false;
    }

    private static TypePropertiesInfo AnalyzeTypeProperties(Compilation compilation, ITypeSymbol classType)
    {
        var propertyNames = new List<string>();
        var unsupportedPropertyNames = new List<string>();

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

            if (p.Type.SpecialType == SpecialType.System_String
                || SupportedPrimitiveTypes.Contains(p.Type.SpecialType)
                || IsSupportedNullableType(compilation, p.Type))
            {
                propertyNames.Add(p.Name);
            }
            else
            {
                unsupportedPropertyNames.Add(p.Name);
            }
        }

        return new TypePropertiesInfo(propertyNames, unsupportedPropertyNames);
    }

    private static bool IsSupportedNullableType(Compilation compilation, ITypeSymbol type)
    {
        if (type.SpecialType != SpecialType.System_Nullable_T)
            return false;

        var nullableT = compilation.GetTypeByMetadataName("System.Nullable`1");

        foreach (var primitiveType in SupportedPrimitiveTypes)
        {
            var nullableType = nullableT?.Construct(compilation.GetSpecialType(primitiveType));
            if (nullableType is null)
                continue;

            if (nullableType.Equals(type, SymbolEqualityComparer.Default))
                return true;
        }

        return false;
    }

    private static readonly SpecialType[] SupportedPrimitiveTypes =
    {
        SpecialType.System_Boolean,
        SpecialType.System_Decimal,
        SpecialType.System_Double,
        SpecialType.System_Int32,
        SpecialType.System_Int64,
        SpecialType.System_Single
    };

    private static void Execute(Compilation compilation, ImmutableArray<ContextClass?> classes, SourceProductionContext context)
    {
        if (classes.IsDefaultOrEmpty)
            return;

        var sb = new StringBuilder();

        foreach (var item in classes)
        {
            if (item is null) continue;

            context.CancellationToken.ThrowIfCancellationRequested();

            sb.Clear();
            GenerateCode(sb, item, compilation, context);
            context.AddSource($"{item.ContextClassType}.g.cs", sb.ToString());
        }
    }

    private static void GenerateHeader(StringBuilder sb)
    {
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("using SpreadCheetah;");
        sb.AppendLine("using SpreadCheetah.SourceGeneration;");
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Buffers;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using System.Threading.Tasks;");
        sb.AppendLine();
    }

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass, Compilation compilation, SourceProductionContext context)
    {
        GenerateHeader(sb);

        var contextType = contextClass.ContextClassType;
        var contextTypeNamespace = contextType.ContainingNamespace;
        if (contextTypeNamespace is { IsGlobalNamespace: false })
            sb.AppendLine($"namespace {contextTypeNamespace}");

        var accessibility = SyntaxFacts.GetText(contextType.DeclaredAccessibility);

        sb.AppendLine("{");
        sb.AppendLine($"    {accessibility} partial class {contextType.Name}");
        sb.AppendLine("    {");
        sb.AppendLine($"        private static {contextType.Name}? _default;");
        sb.AppendLine($"        public static {contextType.Name} Default => _default ??= new {contextType.Name}();");
        sb.AppendLine();
        sb.AppendLine($"        public {contextType.Name}()");
        sb.AppendLine("        {");
        sb.AppendLine("        }");

        var rowTypeNames = new HashSet<string>(StringComparer.Ordinal);

        foreach (var keyValue in contextClass.RowTypes)
        {
            var rowType = keyValue.Key;
            var rowTypeName = rowType.Name;
            if (!rowTypeNames.Add(rowTypeName))
                continue;

            var rowTypeFullName = rowType.ToString();
            var location = keyValue.Value;

            sb.AppendLine();
            sb.AppendLine(2, $"private WorksheetRowTypeInfo<{rowTypeFullName}>? _{rowTypeName};");
            sb.AppendLine(2, $"public WorksheetRowTypeInfo<{rowTypeFullName}> {rowTypeName} => _{rowTypeName} ??= WorksheetRowMetadataServices.CreateObjectInfo<{rowTypeFullName}>(AddAsRowAsync);");
            sb.AppendLine();

            var info = AnalyzeTypeProperties(compilation, rowType);
            ReportDiagnostics(info, rowType, location, contextClass.Options, context);

            GenerateAddAsRow(sb, 2, rowType, info.PropertyNames);

            if (info.PropertyNames.Count > 0)
            {
                sb.AppendLine();
                GenerateAddAsRowInternal(sb, 2, rowTypeFullName, info.PropertyNames);
            }
        }

        sb.AppendLine("    }");
        sb.AppendLine("}");
    }

    private static void ReportDiagnostics(TypePropertiesInfo info, INamedTypeSymbol rowType, Location location, GeneratorOptions? options, SourceProductionContext context)
    {
        if (options?.SuppressWarnings ?? false) return;

        if (info.PropertyNames.Count == 0)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.NoPropertiesFound, location, rowType.Name));

        if (info.UnsupportedPropertyNames.FirstOrDefault() is { } unsupportedProperty)
            context.ReportDiagnostic(Diagnostic.Create(Diagnostics.UnsupportedTypeForCellValue, location, rowType.Name, unsupportedProperty));
    }

    private static void GenerateAddAsRow(StringBuilder sb, int indent, INamedTypeSymbol rowType, List<string> propertyNames)
    {
        sb.AppendIndentation(indent)
            .Append("private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, ")
            .Append(rowType);

        if (rowType.IsReferenceType)
            sb.Append('?');

        sb.AppendLine(" obj, CancellationToken token)");

        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    if (spreadsheet is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(spreadsheet));");

        if (propertyNames.Count == 0)
        {
            sb.AppendLine(indent, "    return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
            sb.AppendLine(indent, "}");
            return;
        }

        if (rowType.IsReferenceType)
        {
            sb.AppendLine(indent + 1, "if (obj is null)");
            sb.AppendLine(indent + 1, "    return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
        }

        sb.AppendLine(indent, "    return AddAsRowInternalAsync(spreadsheet, obj, token);");
        sb.AppendLine(indent, "}");
    }

    private static void GenerateAddAsRowInternal(StringBuilder sb, int indent, string rowTypeFullname, List<string> propertyNames)
    {
        sb.AppendLine(indent, $"private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, {rowTypeFullname} obj, CancellationToken token)");
        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, $"    var cells = ArrayPool<DataCell>.Shared.Rent({propertyNames.Count});");
        sb.AppendLine(indent, "    try");
        sb.AppendLine(indent, "    {");

        for (var i = 0; i < propertyNames.Count; ++i)
        {
            var propertyName = propertyNames[i];
            sb.AppendLine(indent + 2, $"cells[{i}] = new DataCell(obj.{propertyName});");
        }

        sb.AppendLine(indent, $"        await spreadsheet.AddRowAsync(cells.AsMemory(0, {propertyNames.Count}), token).ConfigureAwait(false);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "    finally");
        sb.AppendLine(indent, "    {");
        sb.AppendLine(indent, "        ArrayPool<DataCell>.Shared.Return(cells, true);");
        sb.AppendLine(indent, "    }");
        sb.AppendLine(indent, "}");
    }
}
