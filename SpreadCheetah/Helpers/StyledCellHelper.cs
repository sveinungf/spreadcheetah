namespace SpreadCheetah.Helpers;

internal static class StyledCellHelper
{
    public static ReadOnlySpan<byte> BeginStyledNumberCell => "<c s=\""u8;
    public static ReadOnlySpan<byte> EndReferenceBeginStyleId => "\" s=\""u8;
}
