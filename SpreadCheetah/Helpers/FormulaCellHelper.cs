namespace SpreadCheetah.Helpers;

internal static class FormulaCellHelper
{
    public static ReadOnlySpan<byte> EndStyleBeginFormula => "\"><f>"u8; // TODO: Rename
    public static ReadOnlySpan<byte> BeginNumberFormulaCell => "<c><f>"u8;
    public static ReadOnlySpan<byte> EndFormulaBeginCachedValue => "</f><v>"u8;
    public static ReadOnlySpan<byte> EndCachedValueEndCell => "</v></c>"u8;
}
