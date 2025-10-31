using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using SpreadCheetah.SourceGenerator.Models.Values;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class SymbolExtensions
{
    public static bool IsSupportedType(this ITypeSymbol type)
    {
        return type.IsSupportedValueType()
            || (type.IsNullableType(out var underlyingType) && underlyingType.IsSupportedValueType())
            || type.IsUri();
    }

    private static bool IsSupportedValueType(this ITypeSymbol type)
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
            _ => type.IsFormula()
        };
    }

    public static bool IsFormula(this ITypeSymbol type)
    {
        return type is INamedTypeSymbol
        {
            ContainingNamespace:
            {
                Name: "SpreadCheetah",
                ContainingNamespace.IsGlobalNamespace: true
            },
            IsGenericType: false,
            Name: "Formula",
            TypeKind: TypeKind.Struct
        };
    }

    public static bool IsUri(this ITypeSymbol type)
    {
        return type is INamedTypeSymbol
        {
            ContainingNamespace:
            {
                Name: "System",
                ContainingNamespace.IsGlobalNamespace: true
            },
            IsGenericType: false,
            Name: "Uri",
            TypeKind: TypeKind.Class
        };
    }

    public static PropertyFormula? ToPropertyFormula(this ITypeSymbol type)
    {
        if (type.IsUri())
            return new PropertyFormula(FormulaType.HyperlinkFromUri);
        if (type.IsFormula())
            return new PropertyFormula(FormulaType.GeneralNonNullable);
        if (type.IsNullableType(out var underlyingType) && underlyingType.IsFormula())
            return new PropertyFormula(FormulaType.GeneralNullable);

        return null;
    }

    private static bool IsNullableType(this ITypeSymbol type, [NotNullWhen(true)] out ITypeSymbol? underlyingType)
    {
        if (type is INamedTypeSymbol
            {
                ContainingNamespace:
                {
                    Name: "System",
                    ContainingNamespace.IsGlobalNamespace: true
                },
                IsGenericType: true,
                Name: "Nullable",
                TypeArguments: [ITypeSymbol typeArgument],
                TypeKind: TypeKind.Struct
            })
        {
            underlyingType = typeArgument;
            return true;
        }

        underlyingType = null;
        return false;
    }

    public static bool SupportsCellFormatAttribute(
        this ITypeSymbol type)
    {
        return !type.IsUri() && type.IsSupportedType();
    }

    public static bool IsInstancePropertyWithPublicGetter(
        this IPropertySymbol property)
    {
        return property is
        {
            DeclaredAccessibility: Accessibility.Public,
            IsStatic: false,
            IsIndexer: false,
            IsWriteOnly: false
        };
    }

    public static bool IsStaticPropertyWithGetter(
        this ISymbol symbol,
        [NotNullWhen(true)] out IPropertySymbol? property)
    {
        if (symbol is IPropertySymbol
            {
                GetMethod: not null,
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

    public static bool HasPublicGetter(this IPropertySymbol symbol)
    {
        return symbol is
        {
            DeclaredAccessibility: Accessibility.Public,
            GetMethod.DeclaredAccessibility: Accessibility.Public
        };
    }

    public static IEnumerable<RowProperty> GetClassAndBaseClassProperties(this ITypeSymbol type)
    {
        if ("Object".Equals(type.Name, StringComparison.Ordinal))
            return [];

        InheritedColumnsOrder? inheritedColumnOrder = null;

        foreach (var attribute in type.GetAttributes())
        {
            if (attribute.TryGetInheritedColumnOrderingAttribute(out var order))
            {
                inheritedColumnOrder = order;
                break;
            }
        }

        var inferColumnHeadersInfo = type.GetEffectiveInferColumnHeadersInfo();
        var properties = type
            .GetMembers()
            .OfType<IPropertySymbol>()
            .Select(x => new RowProperty(x, inferColumnHeadersInfo));

        if (inheritedColumnOrder is null || type.BaseType is null)
            return properties;

        var inheritedProperties = GetClassAndBaseClassProperties(type.BaseType);

        return inheritedColumnOrder switch
        {
            InheritedColumnsOrder.InheritedColumnsFirst => [.. inheritedProperties, .. properties],
            InheritedColumnsOrder.InheritedColumnsLast => [.. properties, .. inheritedProperties],
            _ => throw new ArgumentOutOfRangeException(nameof(type), "Unsupported inheritance strategy type")
        };
    }

    public static InferColumnHeadersInfo? GetEffectiveInferColumnHeadersInfo(this ITypeSymbol type)
    {
        foreach (var attribute in type.GetAttributes())
        {
            if (attribute.TryGetInferColumnHeadersAttribute(out var info))
                return info;
        }

        return type.BaseType is { } baseType
            ? baseType.GetEffectiveInferColumnHeadersInfo()
            : null;
    }

    private static bool TryGetInheritedColumnOrderingAttribute(this AttributeData attribute, out InheritedColumnsOrder result)
    {
        result = default;

        if (attribute is not { AttributeClass.MetadataName: Attributes.InheritColumns })
            return false;
        if (!attribute.AttributeClass.HasSpreadCheetahSrcGenNamespace())
            return false;

        result = attribute.NamedArguments is [{ Value.Value: { } arg }] && arg.IsEnum(out InheritedColumnsOrder order)
            ? order
            : InheritedColumnsOrder.InheritedColumnsFirst;

        return true;
    }

    private static bool TryGetInferColumnHeadersAttribute(
        this AttributeData attribute,
        [NotNullWhen(true)] out InferColumnHeadersInfo? result)
    {
        result = null;

        if (attribute is not { AttributeClass.MetadataName: Attributes.InferColumnHeaders })
            return false;
        if (!attribute.AttributeClass.HasSpreadCheetahSrcGenNamespace())
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: INamedTypeSymbol type }])
            return false;

        string? prefix = null;
        string? suffix = null;

        // TODO: Emit warning on invalid Prefix/Suffix values.

        foreach (var (key, value) in attribute.NamedArguments)
        {
            if (value.Value is not string stringValue)
                continue;

            if (string.Equals(key, "Prefix", StringComparison.Ordinal))
                prefix = stringValue;
            else if (string.Equals(key, "Suffix", StringComparison.Ordinal))
                suffix = stringValue;
        }

        result = new InferColumnHeadersInfo(type, prefix, suffix);
        return true;
    }

    public static bool HasCellValueConverterAttribute(this IPropertySymbol property)
    {
        return property.GetAttribute(Attributes.CellValueConverter) is not null;
    }

    public static bool HasColumnIgnoreAttribute(this IPropertySymbol property)
    {
        return property.GetAttribute(Attributes.ColumnIgnore) is not null;
    }

    public static AttributeData? GetColumnOrderAttribute(this IPropertySymbol property)
    {
        return property.GetAttribute(Attributes.ColumnOrder);
    }

    private static AttributeData? GetAttribute(this IPropertySymbol property, string attributeName)
    {
        foreach (var attribute in property.GetAttributes())
        {
            if (attribute.AttributeClass is not { } attributeClass)
                continue;
            if (!string.Equals(attributeName, attributeClass.MetadataName, StringComparison.Ordinal))
                continue;
            if (attributeClass.HasSpreadCheetahSrcGenNamespace())
                return attribute;
        }

        return null;
    }

    public static bool HasSpreadCheetahSrcGenNamespace(this INamedTypeSymbol symbol)
    {
        return symbol is
        {
            ContainingNamespace:
            {
                Name: "SourceGeneration",
                ContainingNamespace:
                {
                    Name: "SpreadCheetah",
                    ContainingNamespace.IsGlobalNamespace: true
                }
            }
        };
    }
}
