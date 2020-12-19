using Microsoft.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace SpreadCheetah.SourceGenerator
{
    internal class ClassPropertiesInfo
    {
        public ITypeSymbol ClassType { get; }
        public List<string> PropertyNames { get; }
        public List<Location> Locations { get; } = new();

        private ClassPropertiesInfo(ITypeSymbol classType, List<string> propertyNames)
        {
            ClassType = classType;
            PropertyNames = propertyNames;
        }

        public static ClassPropertiesInfo CreateFrom(Compilation compilation, ITypeSymbol classType)
        {
            var propertyNames = GetPropertyNames(compilation, classType).ToList();
            return new ClassPropertiesInfo(classType, propertyNames);
        }

        private static IEnumerable<string> GetPropertyNames(Compilation compilation, ITypeSymbol classType)
        {
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

                if (p.Type.SpecialType == SpecialType.System_String)
                    yield return p.Name;

                if (SupportedPrimitiveTypes.Contains(p.Type.SpecialType))
                    yield return p.Name;

                if (IsSupportedNullableType(compilation, p.Type))
                    yield return p.Name;
            }
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
}
