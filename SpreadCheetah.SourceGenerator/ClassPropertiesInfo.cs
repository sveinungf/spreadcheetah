using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator;

internal class ClassPropertiesInfo
{
    public ITypeSymbol ClassType { get; }
    public List<string> PropertyNames { get; }
    public List<string> UnsupportedPropertyNames { get; }
    public List<Location> Locations { get; } = new();

    private ClassPropertiesInfo(ITypeSymbol classType, List<string> propertyNames, List<string> unsupportedPropertyNames)
    {
        ClassType = classType;
        PropertyNames = propertyNames;
        UnsupportedPropertyNames = unsupportedPropertyNames;
    }

    public static ClassPropertiesInfo CreateFrom(Compilation compilation, ITypeSymbol classType)
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

        return new ClassPropertiesInfo(classType, propertyNames, unsupportedPropertyNames);
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
}
