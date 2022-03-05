namespace SpreadCheetah.Validations;

/// <summary>
/// Values that specifies how the user is informed about invalid data in cells with data validation.
/// </summary>
public enum ValidationErrorType
{
    /// <summary>
    /// Blocks the user from entering invalid data.
    /// </summary>
    Blocking,

    /// <summary>
    /// Warns the user if the data is invalid.
    /// </summary>
    Warning,

    /// <summary>
    /// Informs the user if the data is invalid.
    /// </summary>
    Information
}
