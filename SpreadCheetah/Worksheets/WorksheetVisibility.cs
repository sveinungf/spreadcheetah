namespace SpreadCheetah.Worksheets;

/// <summary>
/// Values that specifies the visibility of a worksheet.
/// </summary>
public enum WorksheetVisibility
{
    /// <summary>Visible. The default for worksheets.</summary>
    Visible,

    /// <summary>Hidden. Can be made visible in the Excel UI.</summary>
    Hidden,

    /// <summary>Very hidden. Can not be made visible in the Excel UI.</summary>
    VeryHidden
}
