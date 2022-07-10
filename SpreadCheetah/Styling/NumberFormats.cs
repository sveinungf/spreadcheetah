namespace SpreadCheetah.Styling;

public static class NumberFormats
{
    public const string Fraction = "# ?/?";
    public const string FractionTwoDenominatorPlaces = "# ??/??";
    public const string General = "General";
    public const string NoDecimalPlaces = "0";
    public const string Percent = "0%";
    public const string PercentTwoDecimalPlaces = "0.00%";
    public const string Scientific = "0.00E+00";
    public const string ThousandsSeparator = "#,##0";
    public const string ThousandsSeparatorTwoDecimalPlaces = "#,##0.00";
    public const string TwoDecimalPlaces = "0.00";

    internal static int? GetPredefinedNumberFormatId(string? format) => format switch
    {
        Fraction => 12,
        FractionTwoDenominatorPlaces => 13,
        General => 0,
        NoDecimalPlaces => 1,
        Percent => 9,
        PercentTwoDecimalPlaces => 10,
        Scientific => 11,
        ThousandsSeparator => 3,
        ThousandsSeparatorTwoDecimalPlaces => 4,
        TwoDecimalPlaces => 2,
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
        "@" => 49,
        _ => null
    };
}
