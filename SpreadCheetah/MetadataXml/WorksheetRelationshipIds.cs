namespace SpreadCheetah.MetadataXml;

internal static class WorksheetRelationshipIds
{
    public static ReadOnlySpan<byte> VmlDrawing => "rId1"u8;
    public static ReadOnlySpan<byte> Comments => "rId2"u8;
    public static ReadOnlySpan<byte> Drawing => "rId3"u8;

    public const int TableStartId = 4;
}
