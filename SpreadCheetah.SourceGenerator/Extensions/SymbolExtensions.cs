using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class SymbolExtensions
{
    public static bool IsSupportedType(this ITypeSymbol type)
    {
        return type.SpecialType switch
        {
            SpecialType.System_Boolean => true,
            SpecialType.System_DateTime => true,
            SpecialType.System_Decimal => true,
            SpecialType.System_Double => true,
            SpecialType.System_Int32 => true,
            SpecialType.System_Int64 => true,
            SpecialType.System_Single => true,
            SpecialType.System_String => true,
            _ => type.IsSupportedNullableType(),
        };
    }

    private static bool IsSupportedNullableType(this ITypeSymbol type)
    {
        if (type.NullableAnnotation != NullableAnnotation.Annotated)
            return false;

        return type.ToDisplayString() switch
        {
            "bool?" => true,
            "decimal?" => true,
            "double?" => true,
            "float?" => true,
            "int?" => true,
            "long?" => true,
            "System.DateTime?" => true,
            _ => false,
        };
    }

    public static bool IsInstancePropertyWithPublicGetter(
        this IPropertySymbol property)
    {
        return property is
        {
            DeclaredAccessibility: Accessibility.Public,
            IsStatic: false,
            IsWriteOnly: false
        };
    }

    public static bool IsStaticPropertyWithPublicGetter(
        this ISymbol symbol,
        [NotNullWhen(true)] out IPropertySymbol? property)
    {
        if (symbol is IPropertySymbol
            {
                DeclaredAccessibility: Accessibility.Public,
                GetMethod.DeclaredAccessibility: Accessibility.Public,
                IsStatic: true,
                IsWriteOnly: false
            } p)
        {
            property = p;
            return true;
        }

        property = null;
        return false;
    }

    public static IEnumerable<IPropertySymbol> GetClassAndBaseClassProperties(this ITypeSymbol type)
    {
        if ("Object".Equals(type.Name, StringComparison.Ordinal))
            return [];

        InheritedColumnOrder? inheritedColumnOrder = null;

        foreach (var attribute in type.GetAttributes())
        {
            if (attribute.TryGetInheritedColumnOrderingAttribute(out var order))
            {
                inheritedColumnOrder = order;
                break;
            }
        }

        var properties = type.GetMembers().OfType<IPropertySymbol>();

        if (inheritedColumnOrder is null || type.BaseType is null)
            return properties;

        var inheritedProperties = GetClassAndBaseClassProperties(type.BaseType);

        return inheritedColumnOrder switch
        {
            InheritedColumnOrder.InheritedColumnsFirst => inheritedProperties.Concat(properties),
            InheritedColumnOrder.InheritedColumnsLast => properties.Concat(inheritedProperties),
            _ => throw new ArgumentOutOfRangeException(nameof(type), "Unsupported inheritance strategy type")
        };
    }

    private static bool TryGetInheritedColumnOrderingAttribute(this AttributeData attribute, out InheritedColumnOrder result)
    {
        result = default;

        if (!string.Equals(Attributes.InheritColumns, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return false;

        result = attribute.NamedArguments is [{ Value.Value: { } arg }] && Enum.IsDefined(typeof(InheritedColumnOrder), arg)
            ? (InheritedColumnOrder)arg
            : InheritedColumnOrder.InheritedColumnsFirst;

        return true;
    }

    public static AttributeData? GetCellValueConverterAttribute(this IPropertySymbol property)
    {
        return property
            .GetAttributes()
            .FirstOrDefault(x => Attributes.CellValueConverter.Equals(x.AttributeClass?.ToDisplayString(), StringComparison.Ordinal));
    }

    public static AttributeData? GetColumnOrderAttribute(this IPropertySymbol property)
    {
        return property
            .GetAttributes()
            .FirstOrDefault(x => Attributes.ColumnOrder.Equals(x.AttributeClass?.ToDisplayString(), StringComparison.Ordinal));
    }
}
