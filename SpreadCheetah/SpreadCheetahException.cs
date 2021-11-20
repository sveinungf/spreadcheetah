namespace SpreadCheetah;

/// <summary>
/// The exception that is thrown when <see cref="Spreadsheet" /> is used in an unintended way.
/// </summary>
public class SpreadCheetahException : Exception
{
    /// <inheritdoc/>
    public SpreadCheetahException()
    {
    }

    /// <inheritdoc/>
    public SpreadCheetahException(string message) : base(message)
    {
    }

    /// <inheritdoc/>
    public SpreadCheetahException(string message, Exception exception) : base(message, exception)
    {
    }
}
