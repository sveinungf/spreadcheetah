using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class AttributeDataComparer : IComparer<AttributeData>
{
    public static AttributeDataComparer Instance { get; } = new();

    private AttributeDataComparer()
    {
    }

    public int Compare(AttributeData x, AttributeData y)
    {
        return Compare(x.AttributeClass?.MetadataName, y.AttributeClass?.MetadataName);
    }

    public static int Compare(string? xMetadataName, string? yMetadataName)
    {
        return GetOrder(xMetadataName) - GetOrder(yMetadataName);
    }

    private static int GetOrder(string? metadataName)
    {
        return metadataName switch
        {
            Attributes.ColumnIgnore => -2,
            Attributes.CellValueConverter => -1,
            Attributes.CellStyle => 1,
            _ => 0
        };
    }
}
