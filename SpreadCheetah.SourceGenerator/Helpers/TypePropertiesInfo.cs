namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class TypePropertiesInfo
{
    public List<string> PropertyNames { get; }
    public List<string> UnsupportedPropertyNames { get; }

    public TypePropertiesInfo(List<string> propertyNames, List<string> unsupportedPropertyNames)
    {
        PropertyNames = propertyNames;
        UnsupportedPropertyNames = unsupportedPropertyNames;
    }
}
