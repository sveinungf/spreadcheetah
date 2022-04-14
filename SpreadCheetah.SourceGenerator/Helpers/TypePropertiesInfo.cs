using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class TypePropertiesInfo
{
    public List<string> PropertyNames { get; }
    public List<IPropertySymbol> UnsupportedProperties { get; }

    public TypePropertiesInfo(List<string> propertyNames, List<IPropertySymbol> unsupportedProperties)
    {
        PropertyNames = propertyNames;
        UnsupportedProperties = unsupportedProperties;
    }
}
