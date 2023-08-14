using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class NumberFormatHelper
{
#pragma warning disable CS0618 // Type or member is obsolete - Required for backwards compatibile behaviour
    public static int? GetStandardNumberFormatId(string? format) => format switch
    {
        NumberFormats.General => 0,
        NumberFormats.NoDecimalPlaces => 1,
        NumberFormats.TwoDecimalPlaces => 2,
        NumberFormats.ThousandsSeparator => 3,
        NumberFormats.ThousandsSeparatorTwoDecimalPlaces => 4,
        NumberFormats.Percent => 9,
        NumberFormats.PercentTwoDecimalPlaces => 10,
        NumberFormats.Scientific => 11,
        NumberFormats.Fraction => 12,
        NumberFormats.FractionTwoDenominatorPlaces => 13,
        "mm-dd-yy" => 14,
        "d-mmm-yy" => 15,
        "d-mmm" => 16,
        "mmm-yy" => 17,
        "h:mm AM/PM" => 18,
        "h:mm:ss AM/PM" => 19,
        "h:mm" => 20,
        "h:mm:ss" => 21,
        "m/d/yy h:mm" => 22,
        "#,##0 ;(#,##0)" => 37,
        "#,##0 ;[Red](#,##0)" => 38,
        "#,##0.00;(#,##0.00)" => 39,
        "#,##0.00;[Red](#,##0.00)" => 40,
        "mm:ss" => 45,
        "[h]:mm:ss" => 46,
        "mmss.0" => 47,
        "##0.0E+0" => 48,
        NumberFormats.Text => 49,
        _ => null
    };
#pragma warning restore CS0618 // Type or member is obsolete
}
