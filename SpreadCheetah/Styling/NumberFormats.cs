namespace SpreadCheetah.Styling;

/// <summary>
/// Standard number formats. Can be used in a style by setting <see cref="Style.Format"/>.
/// </summary>
public static class NumberFormats
{
    /// <summary>Format as a sortable date with time.</summary>
    public const string DateTimeSortable = @"yyyy\-mm\-dd\ hh:mm:ss";

#pragma warning disable S1133 // Deprecated code should be removed - This is required for backwards binary compatibility
    /// <summary>Format as integer with fraction using one digit denominator.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.Fraction)})")]
    public const string Fraction = "# ?/?";

    /// <summary>Format as integer with fraction using two digit denominator.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.FractionTwoDenominatorPlaces)})")]
    public const string FractionTwoDenominatorPlaces = "# ??/??";

    /// <summary>The default format.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.General)})")]
    public const string General = "General";

    /// <summary>Format as integer.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.NoDecimalPlaces)})")]
    public const string NoDecimalPlaces = "0";

    /// <summary>Format as integer with percent symbol.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.Percent)})")]
    public const string Percent = "0%";

    /// <summary>Format as number with two decimal places and percent symbol.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.PercentTwoDecimalPlaces)})")]
    public const string PercentTwoDecimalPlaces = "0.00%";

    /// <summary>Format as a number between 1 and 10 multiplied by a power of 10.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.Scientific)})")]
    public const string Scientific = "0.00E+00";

    /// <summary>Format as text.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.Text)})")]
    public const string Text = "@";

    /// <summary>Format as integer with thousands separator.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.ThousandsSeparator)})")]
    public const string ThousandsSeparator = "#,##0";

    /// <summary>Format as number with two decimal places and thousands separator.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.ThousandsSeparatorTwoDecimalPlaces)})")]
    public const string ThousandsSeparatorTwoDecimalPlaces = "#,##0.00";

    /// <summary>Format as number with two decimal places.</summary>
    [Obsolete($"Use {nameof(NumberFormat)}.{nameof(NumberFormat.Standard)}({nameof(StandardNumberFormat)}.{nameof(StandardNumberFormat.TwoDecimalPlaces)})")]
    public const string TwoDecimalPlaces = "0.00";
#pragma warning restore S1133 // Deprecated code should be removed

#pragma warning disable CS0618 // Type or member is obsolete - Required for backwards compatibile behaviour
    internal static int? GetStandardNumberFormatId(string? format) => format switch
    {
        General => 0,
        NoDecimalPlaces => 1,
        TwoDecimalPlaces => 2,
        ThousandsSeparator => 3,
        ThousandsSeparatorTwoDecimalPlaces => 4,
        Percent => 9,
        PercentTwoDecimalPlaces => 10,
        Scientific => 11,
        Fraction => 12,
        FractionTwoDenominatorPlaces => 13,
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
        Text => 49,
        _ => null
    };
#pragma warning restore CS0618 // Type or member is obsolete
}
