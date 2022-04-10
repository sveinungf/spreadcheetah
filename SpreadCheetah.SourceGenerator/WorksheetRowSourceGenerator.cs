using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SpreadCheetah.SourceGenerator.Helpers;
using System.Collections.Immutable;
using System.Text;

namespace SpreadCheetah.SourceGenerators;

[Generator]
public class WorksheetRowSourceGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var filtered = context.SyntaxProvider
            .CreateSyntaxProvider(
                static (s, _) => IsSyntaxTargetForGeneration(s),
                static (ctx, _) => GetSemanticTargetForGeneration(ctx))
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

    private static ContextClass? GetSemanticTargetForGeneration(GeneratorSyntaxContext context)
    {
        if (context.Node is not ClassDeclarationSyntax classDeclaration)
            return null;

        if (!classDeclaration.Modifiers.Any(x => x.IsKind(SyntaxKind.PartialKeyword)))
            return null;

        var classSymbol = context.SemanticModel.GetDeclaredSymbol(classDeclaration);
        if (classSymbol is null)
            return null;

        if (classSymbol.IsStatic)
            return null;

        const string expectedAttributeFullname = "SpreadCheetah.SourceGeneration.WorksheetRowAttribute";
        var expectedAttributeType = context.SemanticModel.Compilation.GetTypeByMetadataName(expectedAttributeFullname);
        if (expectedAttributeType is null)
            return null;

        var rowTypes = new HashSet<INamedTypeSymbol>(SymbolEqualityComparer.Default);

        foreach (var attribute in classSymbol.GetAttributes())
        {
            if (!SymbolEqualityComparer.Default.Equals(expectedAttributeType, attribute.AttributeClass))
                continue;

            var args = attribute.ConstructorArguments;
            if (args.Length != 1)
                continue;

            if (args[0].Value is not INamedTypeSymbol typeSymbol)
                continue;

            rowTypes.Add(typeSymbol);
        }

        return rowTypes.Count > 0
            ? new ContextClass(classSymbol, rowTypes)
            : null;
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

        foreach (var item in classes.Distinct())
        {
            if (item is null) continue;

            context.CancellationToken.ThrowIfCancellationRequested();

            sb.Clear();
            GenerateCode(sb, item, compilation);
            context.AddSource($"{item.ContextClassType}.g.cs", sb.ToString());
        }
    }

    private static void GenerateCode(StringBuilder sb, ContextClass contextClass, Compilation compilation)
    {
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("using SpreadCheetah;");
        sb.AppendLine("using SpreadCheetah.SourceGeneration;");
        sb.AppendLine("using System.Buffers;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using System.Threading.Tasks;");
        sb.AppendLine();

        var contextType = contextClass.ContextClassType;
        var contextTypeNamespace = contextType.ContainingNamespace;
        if (contextTypeNamespace is { IsGlobalNamespace: false })
            sb.AppendLine($"namespace {contextTypeNamespace}");

        var accessibility = SyntaxFacts.GetText(contextType.DeclaredAccessibility);

        sb.AppendLine("{");
        sb.AppendLine($"    {accessibility} partial class {contextType.Name}");
        sb.AppendLine("    {");
        sb.AppendLine($"        private static {contextType.Name}? _default;");
        sb.AppendLine($"        public static {contextType.Name} Default => _default ??= new();");
        sb.AppendLine();
        sb.AppendLine($"        public {contextType.Name}() : base(null)");
        sb.AppendLine("        {");
        sb.AppendLine("        }");

        foreach (var rowType in contextClass.RowTypes)
        {
            var rowTypeName = rowType.Name;
            var rowTypeFullName = rowType.ToString();

            sb.AppendLine();
            sb.AppendLine(2, $"private WorksheetRowTypeInfo<{rowTypeFullName}>? _{rowTypeName};");
            sb.AppendLine(2, $"public WorksheetRowTypeInfo<{rowTypeFullName}> {rowTypeName} => _{rowTypeName} ??= WorksheetRowMetadataServices.CreateObjectInfo<{rowTypeFullName}>(AddAsRowAsync);");
            sb.AppendLine();

            var info = AnalyzeTypeProperties(compilation, rowType);
            GenerateAddAsRow(sb, 2, rowTypeFullName, info.PropertyNames);

            if (info.PropertyNames.Count > 0)
            {
                sb.AppendLine();
                GenerateAddAsRowInternal(sb, 2, rowTypeFullName, info.PropertyNames);
            }
        }

        sb.AppendLine("    }");
        sb.AppendLine("}");
    }

    private static void GenerateAddAsRow(StringBuilder sb, int indent, string rowTypeFullname, List<string> propertyNames)
    {
        sb.AppendLine(indent, $"private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, {rowTypeFullname}? obj, CancellationToken token)");
        sb.AppendLine(indent, "{");
        sb.AppendLine(indent, "    if (spreadsheet is null)");
        sb.AppendLine(indent, "        throw new ArgumentNullException(nameof(spreadsheet));");

        if (propertyNames.Count == 0)
        {
            sb.AppendLine(indent + 1, "return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
        }
        else
        {
            sb.AppendLine(indent + 1, "if (obj is null)");
            sb.AppendLine(indent + 1, "    return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
            sb.AppendLine(indent + 1, "return AddAsRowInternalAsync(spreadsheet, obj, token);");
        }

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
