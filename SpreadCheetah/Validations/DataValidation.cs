using SpreadCheetah.Helpers;
using System.Globalization;

namespace SpreadCheetah.Validations;

public sealed class DataValidation
{
    private const string MinGreaterThanMaxMessage = "The min value must be less than or equal to the max value.";

    internal ValidationType Type { get; }
    internal ValidationOperator Operator { get; }
    internal string? Value1 { get; }
    internal string? Value2 { get; }

    private DataValidation(
        ValidationType type,
        ValidationOperator op,
        string? value1,
        string? value2 = null)
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

    private static DataValidation Decimal(ValidationOperator op, double value1, double? value2 = null)
    {
        return new DataValidation(ValidationType.Decimal, op,
            value1.ToString(CultureInfo.InvariantCulture),
            value2?.ToString(CultureInfo.InvariantCulture));
    }

    public static DataValidation DecimalBetween(double min, double max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : Decimal(ValidationOperator.Between, min, max);

    public static DataValidation DecimalNotBetween(double min, double max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : Decimal(ValidationOperator.NotBetween, min, max);

    public static DataValidation DecimalEqualTo(double value) => Decimal(ValidationOperator.EqualTo, value);
    public static DataValidation DecimalNotEqualTo(double value) => Decimal(ValidationOperator.NotEqualTo, value);
    public static DataValidation DecimalGreaterThan(double value) => Decimal(ValidationOperator.GreaterThan, value);
    public static DataValidation DecimalGreaterThanOrEqualTo(double value) => Decimal(ValidationOperator.GreaterThanOrEqualTo, value);
    public static DataValidation DecimalLessThan(double value) => Decimal(ValidationOperator.LessThan, value);
    public static DataValidation DecimalLessThanOrEqualTo(double value) => Decimal(ValidationOperator.LessThanOrEqualTo, value);

    private static DataValidation Integer(ValidationOperator op, int value1, int? value2 = null)
    {
        return new DataValidation(ValidationType.Integer, op,
            value1.ToString(CultureInfo.InvariantCulture),
            value2?.ToString(CultureInfo.InvariantCulture));
    }

    public static DataValidation IntegerBetween(int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : Integer(ValidationOperator.Between, min, max);

    public static DataValidation IntegerNotBetween(int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : Integer(ValidationOperator.NotBetween, min, max);

    public static DataValidation IntegerEqualTo(int value) => Integer(ValidationOperator.EqualTo, value);
    public static DataValidation IntegerNotEqualTo(int value) => Integer(ValidationOperator.NotEqualTo, value);
    public static DataValidation IntegerGreaterThan(int value) => Integer(ValidationOperator.GreaterThan, value);
    public static DataValidation IntegerGreaterThanOrEqualTo(int value) => Integer(ValidationOperator.GreaterThanOrEqualTo, value);
    public static DataValidation IntegerLessThan(int value) => Integer(ValidationOperator.LessThan, value);
    public static DataValidation IntegerLessThanOrEqualTo(int value) => Integer(ValidationOperator.LessThanOrEqualTo, value);
}
