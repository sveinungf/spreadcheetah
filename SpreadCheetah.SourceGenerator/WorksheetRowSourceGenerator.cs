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

    private static TypePropertiesInfo CreateFrom(Compilation compilation, ITypeSymbol classType)
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

        foreach (var item in classes.Distinct())
        {
            if (item is null) continue;

            context.CancellationToken.ThrowIfCancellationRequested();

            var rowType = item.RowTypes.First(); // TODO

            var info = CreateFrom(compilation, rowType);
            var dict = new Dictionary<ITypeSymbol, TypePropertiesInfo>(SymbolEqualityComparer.Default)
            {
                { rowType, info }
            };

            var sb = new StringBuilder();
            GenerateCode(sb, item.ContextClassType, dict);
            context.AddSource($"{item.ContextClassType}.g.cs", sb.ToString());
        }
    }

    private static void GenerateCode(StringBuilder sb, ITypeSymbol contextType, Dictionary<ITypeSymbol, TypePropertiesInfo> typePropertiesInfos)
    {
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("using SpreadCheetah;");
        sb.AppendLine("using SpreadCheetah.SourceGeneration;");
        sb.AppendLine("using System.Buffers;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using System.Threading.Tasks;");
        sb.AppendLine();

        var rowType = typePropertiesInfos.First(); // TODO
        var rowTypeName = rowType.Key.Name;
        var rowTypeFullName = rowType.Key.ToString();
        var info = rowType.Value;

        var contextTypeNamespace = contextType.ContainingNamespace;
        if (contextTypeNamespace is { IsGlobalNamespace: false })
            sb.AppendLine($"namespace {contextTypeNamespace}");

        sb.AppendLine("{");
        sb.AppendLine($"    public partial class {contextType.Name}");
        sb.AppendLine("    {");
        sb.AppendLine($"        private static {contextType.Name}? _default;");
        sb.AppendLine($"        public static {contextType.Name} Default => _default ??= new();");
        sb.AppendLine();
        sb.AppendLine($"        private WorksheetRowTypeInfo<{rowTypeFullName}>? _{rowTypeName};");
        sb.AppendLine($"        public WorksheetRowTypeInfo<{rowTypeFullName}> {rowTypeName} => _{rowTypeName} ??= WorksheetRowMetadataServices.CreateObjectInfo<{rowTypeFullName}>(AddAsRowAsync);");
        sb.AppendLine();
        sb.AppendLine($"        public {contextType.Name}() : base(null)");
        sb.AppendLine("        {");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.AppendLine($"        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, {rowTypeFullName}? obj, CancellationToken token)");
        sb.AppendLine("        {");
        sb.AppendLine("            if (spreadsheet is null)");
        sb.AppendLine("                throw new ArgumentNullException(nameof(spreadsheet));");

        if (info.PropertyNames.Count == 0)
        {
            sb.AppendLine(3, "return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
        }
        else
        {
            sb.AppendLine(3, "if (obj is null)");
            sb.AppendLine(3, "    return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);");
            sb.AppendLine(3, "return AddAsRowInternalAsync(spreadsheet, obj, token);");
        }

        sb.AppendLine("        }");

        if (info.PropertyNames.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine(2, $"private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, {rowTypeFullName} obj, CancellationToken token)");
            sb.AppendLine(2, "{");
            sb.AppendLine(2, $"    var cells = ArrayPool<DataCell>.Shared.Rent({info.PropertyNames.Count});");
            sb.AppendLine(2, "    try");
            sb.AppendLine(2, "    {");

            var cellCount = info.PropertyNames.Count;
            for (var i = 0; i < cellCount; ++i)
            {
                var propertyName = info.PropertyNames[i];
                sb.AppendLine(4, $"cells[{i}] = new DataCell(obj.{propertyName});");
            }

            sb.AppendLine(2, $"        await spreadsheet.AddRowAsync(cells.AsMemory(0, {cellCount}), token).ConfigureAwait(false);");
            sb.AppendLine(2, "    }");
            sb.AppendLine(2, "    finally");
            sb.AppendLine(2, "    {");
            sb.AppendLine(2, "        ArrayPool<DataCell>.Shared.Return(cells, true);");
            sb.AppendLine(2, "    }");
            sb.AppendLine(2, "}");
        }

        sb.AppendLine("    }");
        sb.AppendLine("}");
    }
}
