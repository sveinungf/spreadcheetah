using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
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

    public static bool IsPropertyWithPublicGetter(
        this ISymbol symbol,
        [NotNullWhen(true)] out IPropertySymbol? property)
    {
        if (symbol is IPropertySymbol
            {
                DeclaredAccessibility: Accessibility.Public,
                IsStatic: false,
                IsWriteOnly: false
            } p)
        {
            property = p;
            return true;
        }

        property = null;
        return false;
    }

    public static RowTypeProperty ToRowTypeProperty(
        this IPropertySymbol p,
        TypedConstant? columnHeaderAttributeValue)
    {
        var columnHeader = columnHeaderAttributeValue?.ToCSharpString() ?? @$"""{p.Name}""";

        return new RowTypeProperty(
            ColumnHeader: columnHeader,
            Name: p.Name,
            TypeFullName: p.Type.ToDisplayString(),
            TypeName: p.Type.Name,
            TypeNullableAnnotation: p.NullableAnnotation,
            TypeSpecialType: p.Type.SpecialType);
    }
}
