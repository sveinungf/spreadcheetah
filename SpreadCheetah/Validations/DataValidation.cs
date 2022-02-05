using SpreadCheetah.Helpers;
using System.Globalization;

namespace SpreadCheetah.Validations;

public sealed class DataValidation
{
    internal ValidationType Type { get; }
    internal ValidationOperator Operator { get; }
    internal string? Value1 { get; }
    internal string? Value2 { get; }

    private DataValidation(
        ValidationType type,
        ValidationOperator op,
        string? value1,
        string? value2)
    {
        Type = type;
        Operator = op;
        Value1 = value1;
        Value2 = value2;
    }

    public bool IgnoreBlank { get; set; } = true;
    public bool ShowErrorAlert { get; set; } = true;
    public bool ShowInputMessage { get; set; } = true;
    public string? ErrorTitle { get; set; } // TODO: Validate length
    public string? ErrorMessage { get; set; } // TODO: Validate length
    public string? InputTitle { get; set; } // TODO: Validate length
    public string? InputMessage { get; set; } // TODO: Validate length

    public ValidationErrorType ErrorType
    {
        get => _errorType;
        set => _errorType = EnumHelper.IsDefined(value)
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), value, null);
    }

    private ValidationErrorType _errorType;

    public static DataValidation IntegerBetween(int minimum, int maximum)
    {
        // TODO: Validate arguments
        return new DataValidation(
            ValidationType.Integer,
            ValidationOperator.Between,
            minimum.ToString(CultureInfo.InvariantCulture),
            maximum.ToString(CultureInfo.InvariantCulture));
    }

    public static DataValidation IntegerNotBetween(int minimum, int maximum)
    {
        // TODO: Validate arguments
        return new DataValidation(
            ValidationType.Integer,
            ValidationOperator.NotBetween,
            minimum.ToString(CultureInfo.InvariantCulture),
            maximum.ToString(CultureInfo.InvariantCulture));
    }
}
