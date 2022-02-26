using SpreadCheetah.Helpers;
using System.Net;
using System.Text;

namespace SpreadCheetah.Validations;

public sealed class DataValidation
{
    private const string MinGreaterThanMaxMessage = "The min value must be less than or equal to the max value.";

    internal ValidationType Type { get; }
    internal ValidationOperator Operator { get; }
    internal string? Value1 { get; }
    internal string? Value2 { get; }
    internal bool ShowDropdown { get; }

    private DataValidation(
        ValidationType type,
        ValidationOperator op,
        string? value1,
        string? value2 = null,
        bool showDropdown = true)
    {
        Type = type;
        Operator = op;
        ShowDropdown = showDropdown;
        Value1 = value1;
        Value2 = value2;
    }

    public bool IgnoreBlank { get; set; } = true;
    public bool ShowErrorAlert { get; set; } = true;
    public bool ShowInputMessage { get; set; } = true;

    public string? ErrorTitle { get => _errorTitle; set => _errorTitle = VerifyLength(value, 32); }
    private string? _errorTitle;

    public string? ErrorMessage { get => _errorMessage; set => _errorMessage = VerifyLength(value, 255); }
    private string? _errorMessage;

    public string? InputTitle { get => _inputTitle; set => _inputTitle = VerifyLength(value, 32); }
    private string? _inputTitle;

    public string? InputMessage { get => _inputMessage; set => _inputMessage = VerifyLength(value, 255); }
    private string? _inputMessage;

    private static string? VerifyLength(string? value, int maxLength) =>
        value is not null && value.Length > maxLength
            ? throw new ArgumentException($"The value can not exceed {maxLength} characters.", nameof(value))
            : value;

    public ValidationErrorType ErrorType
    {
        get => _errorType;
        set => _errorType = EnumHelper.IsDefined(value)
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), value, null);
    }

    private ValidationErrorType _errorType;

    private static DataValidation Decimal(ValidationOperator op, double min, double max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : new DataValidation(ValidationType.Decimal, op, min.ToStringInvariant(), max.ToStringInvariant());

    private static DataValidation Decimal(ValidationOperator op, double value) => new(ValidationType.Decimal, op, value.ToStringInvariant());
    public static DataValidation DecimalBetween(double min, double max) => Decimal(ValidationOperator.Between, min, max);
    public static DataValidation DecimalNotBetween(double min, double max) => Decimal(ValidationOperator.NotBetween, min, max);
    public static DataValidation DecimalEqualTo(double value) => Decimal(ValidationOperator.EqualTo, value);
    public static DataValidation DecimalNotEqualTo(double value) => Decimal(ValidationOperator.NotEqualTo, value);
    public static DataValidation DecimalGreaterThan(double value) => Decimal(ValidationOperator.GreaterThan, value);
    public static DataValidation DecimalGreaterThanOrEqualTo(double value) => Decimal(ValidationOperator.GreaterThanOrEqualTo, value);
    public static DataValidation DecimalLessThan(double value) => Decimal(ValidationOperator.LessThan, value);
    public static DataValidation DecimalLessThanOrEqualTo(double value) => Decimal(ValidationOperator.LessThanOrEqualTo, value);

    private static DataValidation Integer(ValidationOperator op, int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : new DataValidation(ValidationType.Integer, op, min.ToStringInvariant(), max.ToStringInvariant());

    private static DataValidation Integer(ValidationOperator op, int value) => new(ValidationType.Integer, op, value.ToStringInvariant());
    public static DataValidation IntegerBetween(int min, int max) => Integer(ValidationOperator.Between, min, max);
    public static DataValidation IntegerNotBetween(int min, int max) => Integer(ValidationOperator.NotBetween, min, max);
    public static DataValidation IntegerEqualTo(int value) => Integer(ValidationOperator.EqualTo, value);
    public static DataValidation IntegerNotEqualTo(int value) => Integer(ValidationOperator.NotEqualTo, value);
    public static DataValidation IntegerGreaterThan(int value) => Integer(ValidationOperator.GreaterThan, value);
    public static DataValidation IntegerGreaterThanOrEqualTo(int value) => Integer(ValidationOperator.GreaterThanOrEqualTo, value);
    public static DataValidation IntegerLessThan(int value) => Integer(ValidationOperator.LessThan, value);
    public static DataValidation IntegerLessThanOrEqualTo(int value) => Integer(ValidationOperator.LessThanOrEqualTo, value);

    private static DataValidation TextLength(ValidationOperator op, int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : new DataValidation(ValidationType.TextLength, op, min.ToStringInvariant(), max.ToStringInvariant());

    private static DataValidation TextLength(ValidationOperator op, int value) => new(ValidationType.TextLength, op, value.ToStringInvariant());
    public static DataValidation TextLengthBetween(int min, int max) => TextLength(ValidationOperator.Between, min, max);
    public static DataValidation TextLengthNotBetween(int min, int max) => TextLength(ValidationOperator.NotBetween, min, max);
    public static DataValidation TextLengthEqualTo(int value) => TextLength(ValidationOperator.EqualTo, value);
    public static DataValidation TextLengthNotEqualTo(int value) => TextLength(ValidationOperator.NotEqualTo, value);
    public static DataValidation TextLengthGreaterThan(int value) => TextLength(ValidationOperator.GreaterThan, value);
    public static DataValidation TextLengthGreaterThanOrEqualTo(int value) => TextLength(ValidationOperator.GreaterThanOrEqualTo, value);
    public static DataValidation TextLengthLessThan(int value) => TextLength(ValidationOperator.LessThan, value);
    public static DataValidation TextLengthLessThanOrEqualTo(int value) => TextLength(ValidationOperator.LessThanOrEqualTo, value);

    public static DataValidation ListValues(IEnumerable<string> values, bool showDropdown = true)
    {
        if (values is null)
            throw new ArgumentNullException(nameof(values));

        var sb = new StringBuilder();
        sb.Append('"');
        var first = true;

        foreach (var value in values)
        {
            if (value.ContainsChar(','))
                throw new ArgumentException($"Commas are not allowed in the list values. This value contains a comma: {value}", nameof(values));

            if (!first)
                sb.Append(',');

            sb.Append(WebUtility.HtmlEncode(value));
            first = false;
        }

        sb.Append('"');

        return new DataValidation(ValidationType.List, ValidationOperator.None, sb.ToString(), null, showDropdown);
    }
}
