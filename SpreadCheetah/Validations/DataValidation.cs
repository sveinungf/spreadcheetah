using SpreadCheetah.Helpers;
using System.Net;
using System.Text;

namespace SpreadCheetah.Validations;

/// <summary>
/// Represents data validation for one or more worksheet cells.
/// </summary>
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

    /// <summary>Ignore cells without values. Defaults to <c>true</c>.</summary>
    public bool IgnoreBlank { get; set; } = true;
    /// <summary>Show an error alert for invalid values. Defaults to <c>true</c>.</summary>
    public bool ShowErrorAlert { get; set; } = true;
    /// <summary>Show the input message box. Defaults to <c>true</c>.</summary>
    public bool ShowInputMessage { get; set; } = true;

    /// <summary>Title for error alerts. Maximum 32 characters.</summary>
    public string? ErrorTitle { get => _errorTitle; set => _errorTitle = value.WithEnsuredMaxLength(32); }
    private string? _errorTitle;

    /// <summary>Message for error alerts. Maximum 255 characters.</summary>
    public string? ErrorMessage { get => _errorMessage; set => _errorMessage = value.WithEnsuredMaxLength(255); }
    private string? _errorMessage;

    /// <summary>Title for the input message box. Maximum 32 characters.</summary>
    public string? InputTitle { get => _inputTitle; set => _inputTitle = value.WithEnsuredMaxLength(32); }
    private string? _inputTitle;

    /// <summary>Message for the input message box. Maximum 255 characters.</summary>
    public string? InputMessage { get => _inputMessage; set => _inputMessage = value.WithEnsuredMaxLength(255); }
    private string? _inputMessage;

    /// <summary>
    /// Specify how the user is informed about cells with invalid data.
    /// Defaults to <see cref="ValidationErrorType.Blocking"/>.
    /// </summary>
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

    /// <summary>Validate that decimals are between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation DecimalBetween(double min, double max) => Decimal(ValidationOperator.Between, min, max);
    /// <summary>Validate that decimals are not between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation DecimalNotBetween(double min, double max) => Decimal(ValidationOperator.NotBetween, min, max);
    /// <summary>Validate that decimals are equal to <paramref name="value"/>.</summary>
    public static DataValidation DecimalEqualTo(double value) => Decimal(ValidationOperator.EqualTo, value);
    /// <summary>Validate that decimals are not equal to <paramref name="value"/>.</summary>
    public static DataValidation DecimalNotEqualTo(double value) => Decimal(ValidationOperator.NotEqualTo, value);
    /// <summary>Validate that decimals are greater than <paramref name="value"/>.</summary>
    public static DataValidation DecimalGreaterThan(double value) => Decimal(ValidationOperator.GreaterThan, value);
    /// <summary>Validate that decimals are greater than or equal to <paramref name="value"/>.</summary>
    public static DataValidation DecimalGreaterThanOrEqualTo(double value) => Decimal(ValidationOperator.GreaterThanOrEqualTo, value);
    /// <summary>Validate that decimals are less than <paramref name="value"/>.</summary>
    public static DataValidation DecimalLessThan(double value) => Decimal(ValidationOperator.LessThan, value);
    /// <summary>Validate that decimals are less than or equal to <paramref name="value"/>.</summary>
    public static DataValidation DecimalLessThanOrEqualTo(double value) => Decimal(ValidationOperator.LessThanOrEqualTo, value);

    private static DataValidation Integer(ValidationOperator op, int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : new DataValidation(ValidationType.Integer, op, min.ToStringInvariant(), max.ToStringInvariant());

    private static DataValidation Integer(ValidationOperator op, int value) => new(ValidationType.Integer, op, value.ToStringInvariant());
    /// <summary>Validate that integers are between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation IntegerBetween(int min, int max) => Integer(ValidationOperator.Between, min, max);
    /// <summary>Validate that integers are not between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation IntegerNotBetween(int min, int max) => Integer(ValidationOperator.NotBetween, min, max);
    /// <summary>Validate that integers are equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerEqualTo(int value) => Integer(ValidationOperator.EqualTo, value);
    /// <summary>Validate that integers are not equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerNotEqualTo(int value) => Integer(ValidationOperator.NotEqualTo, value);
    /// <summary>Validate that integers are greater than <paramref name="value"/>.</summary>
    public static DataValidation IntegerGreaterThan(int value) => Integer(ValidationOperator.GreaterThan, value);
    /// <summary>Validate that integers are greater than or equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerGreaterThanOrEqualTo(int value) => Integer(ValidationOperator.GreaterThanOrEqualTo, value);
    /// <summary>Validate that integers are less than <paramref name="value"/>.</summary>
    public static DataValidation IntegerLessThan(int value) => Integer(ValidationOperator.LessThan, value);
    /// <summary>Validate that integers are less than or equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerLessThanOrEqualTo(int value) => Integer(ValidationOperator.LessThanOrEqualTo, value);

    private static DataValidation TextLength(ValidationOperator op, int min, int max) => max < min
        ? throw new ArgumentException(MinGreaterThanMaxMessage, nameof(min))
        : new DataValidation(ValidationType.TextLength, op, min.ToStringInvariant(), max.ToStringInvariant());

    private static DataValidation TextLength(ValidationOperator op, int value) => new(ValidationType.TextLength, op, value.ToStringInvariant());
    /// <summary>Validate that text lengths are between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation TextLengthBetween(int min, int max) => TextLength(ValidationOperator.Between, min, max);
    /// <summary>Validate that text length are not between <paramref name="min"/> and <paramref name="max"/>.</summary>
    public static DataValidation TextLengthNotBetween(int min, int max) => TextLength(ValidationOperator.NotBetween, min, max);
    /// <summary>Validate that text lengths are equal to <paramref name="value"/>.</summary>
    public static DataValidation TextLengthEqualTo(int value) => TextLength(ValidationOperator.EqualTo, value);
    /// <summary>Validate that text lengths are not equal to <paramref name="value"/>.</summary>
    public static DataValidation TextLengthNotEqualTo(int value) => TextLength(ValidationOperator.NotEqualTo, value);
    /// <summary>Validate that text lengths are greater than <paramref name="value"/>.</summary>
    public static DataValidation TextLengthGreaterThan(int value) => TextLength(ValidationOperator.GreaterThan, value);
    /// <summary>Validate that text lengths are greater than or equal to <paramref name="value"/>.</summary>
    public static DataValidation TextLengthGreaterThanOrEqualTo(int value) => TextLength(ValidationOperator.GreaterThanOrEqualTo, value);
    /// <summary>Validate that text lengths are less than <paramref name="value"/>.</summary>
    public static DataValidation TextLengthLessThan(int value) => TextLength(ValidationOperator.LessThan, value);
    /// <summary>Validate that text lengths are less than or equal to <paramref name="value"/>.</summary>
    public static DataValidation TextLengthLessThanOrEqualTo(int value) => TextLength(ValidationOperator.LessThanOrEqualTo, value);

    /// <summary>Validate that cell values equal any of <paramref name="values"/>.</summary>
    public static DataValidation ListValues(IEnumerable<string> values, bool showDropdown = true)
    {
        if (values is null)
            throw new ArgumentNullException(nameof(values));

        var sb = new StringBuilder();
        sb.Append('"');
        var first = true;

        foreach (var value in values)
        {
            if (value.Contains(',', StringComparison.Ordinal))
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
