using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Text;

#if !NET6_0_OR_GREATER
using ArgumentNullException = SpreadCheetah.Helpers.Backporting.ArgumentNullExceptionBackport;
#endif

namespace SpreadCheetah.Validations;

/// <summary>
/// Represents data validation for one or more worksheet cells.
/// </summary>
public sealed class DataValidation
{
    private const int MaxValueLength = 255;
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

    /// <summary>Ignore cells without values. Defaults to <see langword="true"/>.</summary>
    public bool IgnoreBlank { get; set; } = true;
    /// <summary>Show an error alert for invalid values. Defaults to <see langword="true"/>.</summary>
    public bool ShowErrorAlert { get; set; } = true;
    /// <summary>Show the input message box. Defaults to <see langword="true"/>.</summary>
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
        set => _errorType = EnumPolyfill.IsDefined(value)
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

    private static DataValidation ListValuesInternal(string value, bool showDropdown) => new(ValidationType.List, ValidationOperator.None, value, null, showDropdown);

    /// <summary>
    /// Validate that cell values equal any of the values from the specified cell range.
    /// The cell range refers to cells within the same worksheet as the data validation.
    /// To reference cells in another worksheet, use the overload with the worksheet name parameter.
    /// The cell range must be in the A1 reference style. Some examples:
    /// <list type="bullet">
    ///   <item><term><c>A1:A4</c></term><description>References the range from cell A1 to A4.</description></item>
    ///   <item><term><c>A1:E5</c></term><description>References the range from cell A1 to E5.</description></item>
    ///   <item><term><c>A1:A1048576</c></term><description>References all cells in column A.</description></item>
    ///   <item><term><c>A5:XFD5</c></term><description>References all cells in row 5.</description></item>
    /// </list>
    /// </summary>
    public static DataValidation ListValuesFromCells(string cellRange, bool showDropdown = true)
    {
        ArgumentNullException.ThrowIfNull(cellRange);
        var cellReference = SingleCellOrCellRangeReference.Create(cellRange);
        return ListValuesInternal(cellReference.Reference, showDropdown);
    }

    /// <summary>
    /// Validate that cell values equal any of the values from the specified cell range.
    /// The cell range refers to cells within the specified worksheet.
    /// To reference cells in the same worksheet as the data validation, you can use the overload without the worksheet name parameter.
    /// The cell range must be in the A1 reference style. Some examples:
    /// <list type="bullet">
    ///   <item><term><c>A1:A4</c></term><description>References the range from cell A1 to A4.</description></item>
    ///   <item><term><c>A1:E5</c></term><description>References the range from cell A1 to E5.</description></item>
    ///   <item><term><c>A1:A1048576</c></term><description>References all cells in column A.</description></item>
    ///   <item><term><c>A5:XFD5</c></term><description>References all cells in row 5.</description></item>
    /// </list>
    /// </summary>
    public static DataValidation ListValuesFromCells(string worksheetName, string cellRange, bool showDropdown = true)
    {
        ArgumentNullException.ThrowIfNull(worksheetName);
        ArgumentNullException.ThrowIfNull(cellRange);
        worksheetName.EnsureValidWorksheetName();

        var cellReference = SingleCellOrCellRangeReference.Create(cellRange);

#pragma warning disable CA1307, MA0074 // Specify StringComparison for clarity
        worksheetName = worksheetName.Replace("'", "''");
#pragma warning restore CA1307, MA0074 // Specify StringComparison for clarity

        var value = $"&apos;{XmlUtility.XmlEncode(worksheetName)}&apos;!{cellReference.Reference}";
        return ListValuesInternal(value, showDropdown);
    }

    /// <summary>
    /// Validate that cell values equal any of <paramref name="values"/>.
    /// Requirements:
    /// <list type="bullet">
    ///   <item><description>The values can't contain commas.</description></item>
    ///   <item><description>The combined length of the values (including required comma separators) can't exceed 255 characters.</description></item>
    /// </list>
    /// An <see cref="ArgumentException"/> will be thrown if any of these requirements are not met.
    /// </summary>
    public static DataValidation ListValues(IEnumerable<string> values, bool showDropdown = true)
    {
        ArgumentNullException.ThrowIfNull(values);

        if (TryCreateListValuesInternal(values, showDropdown, out var invalidValue, out var dataValidation))
            return dataValidation;

        if (invalidValue is not null)
            throw new ArgumentException($"Commas are not allowed in the list values. This value contains a comma: {invalidValue}", nameof(values));

        throw new ArgumentException($"The combined length of the values (including required comma separators) can't exceed {MaxValueLength} characters.", nameof(values));
    }

    /// <summary>
    /// Validate that cell values equal any of <paramref name="values"/>.
    /// Requirements:
    /// <list type="bullet">
    ///   <item><description>The values can't contain commas.</description></item>
    ///   <item><description>The combined length of the values (including required comma separators) can't exceed 255 characters.</description></item>
    /// </list>
    /// Returns <see langword="false"/> if any of these requirements are not met.
    /// </summary>
    public static bool TryCreateListValues(IEnumerable<string> values, bool showDropdown, [NotNullWhen(true)] out DataValidation? dataValidation)
    {
        ArgumentNullException.ThrowIfNull(values);
        return TryCreateListValuesInternal(values, showDropdown, out _, out dataValidation);
    }

    private static bool TryCreateListValuesInternal(
        IEnumerable<string> values,
        bool showDropdown,
        out string? invalidValue,
        [NotNullWhen(true)] out DataValidation? dataValidation)
    {
        invalidValue = null;
        dataValidation = null;
        var sb = new StringBuilder();
        sb.Append('"');
        var first = true;
        int combinedLength = 0;

        foreach (var value in values)
        {
            if (value.Contains(',', StringComparison.Ordinal))
            {
                invalidValue = value;
                return false;
            }

            if (!first)
            {
                sb.Append(',');
                ++combinedLength;
            }

            sb.Append(XmlUtility.XmlEncode(value));
            first = false;

            // Character length (and not the encoded character length) is used to calculate the combined length.
            combinedLength += value.Length;
            if (combinedLength > MaxValueLength)
                return false;
        }

        sb.Append('"');

        dataValidation = ListValuesInternal(sb.ToString(), showDropdown);
        return true;
    }
}
