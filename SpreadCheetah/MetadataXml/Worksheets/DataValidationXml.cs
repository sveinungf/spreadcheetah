using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml.Worksheets;

internal struct DataValidationXml(
    SingleCellOrCellRangeReference reference,
    DataValidation validation,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private BufferWriteProgress _progress;

#pragma warning disable EPS12 // A struct member can be made readonly
    public bool TryWrite()
#pragma warning restore EPS12 // A struct member can be made readonly
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        var valid = validation;

        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.InputTitle => TryWriteAttribute(valid.InputTitle, " promptTitle=\""u8),
            Element.InputMessage => TryWriteAttribute(valid.InputMessage, " prompt=\""u8),
            Element.ErrorTitle => TryWriteAttribute(valid.ErrorTitle, " errorTitle=\""u8),
            Element.ErrorMessage => TryWriteAttribute(valid.ErrorMessage, " error=\""u8),
            Element.Reference => TryWriteReference(),
            Element.Value1 => TryWriteValue1(),
            Element.Value2 => TryWriteValue2(),
            _ => buffer.TryWrite("</dataValidation>"u8)
        };

        if (Current)
        {
            _progress = default;
            ++_next;
        }

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var valid = validation;

        var type = new SpanByteAttribute("type"u8, GetTypeValue(valid.Type));
        var errorType = new SpanByteAttribute("errorStyle"u8, GetErrorTypeValue(valid.ErrorType));
        var theOperator = new SpanByteAttribute("operator"u8, GetOperatorValue(valid.Operator));
        var allowBlank = new BooleanAttribute("allowBlank"u8, valid.IgnoreBlank ? true : null);
        var showDropdown = new BooleanAttribute("showDropDown"u8, valid.ShowDropdown ? null : true);
        var showInputMessage = new BooleanAttribute("showInputMessage"u8, valid.ShowInputMessage ? true : null);
        var showErrorMessage = new BooleanAttribute("showErrorMessage"u8, valid.ShowErrorAlert ? true : null);

        return buffer.TryWrite(
            $"{"<dataValidation"u8}" +
            $"{type}" +
            $"{errorType}" +
            $"{theOperator}" +
            $"{allowBlank}" +
            $"{showDropdown}" +
            $"{showInputMessage}" +
            $"{showErrorMessage}");
    }

    private static ReadOnlySpan<byte> GetTypeValue(ValidationType type) => type switch
    {
        ValidationType.DateTime => "date"u8,
        ValidationType.Decimal => "decimal"u8,
        ValidationType.Integer => "whole"u8,
        ValidationType.List => "list"u8,
        _ => "textLength"u8
    };

    private static ReadOnlySpan<byte> GetErrorTypeValue(ValidationErrorType type) => type switch
    {
        ValidationErrorType.Warning => "warning"u8,
        ValidationErrorType.Information => "information"u8,
        _ => []
    };

    private static ReadOnlySpan<byte> GetOperatorValue(ValidationOperator op) => op switch
    {
        ValidationOperator.NotBetween => "notBetween"u8,
        ValidationOperator.EqualTo => "equal"u8,
        ValidationOperator.NotEqualTo => "notEqual"u8,
        ValidationOperator.GreaterThan => "greaterThan"u8,
        ValidationOperator.LessThan => "lessThan"u8,
        ValidationOperator.GreaterThanOrEqualTo => "greaterThanOrEqual"u8,
        ValidationOperator.LessThanOrEqualTo => "lessThanOrEqual"u8,
        _ => []
    };

    private bool TryWriteAttribute(string? value, scoped ReadOnlySpan<byte> attributeName)
    {
        if (string.IsNullOrEmpty(value))
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{attributeName}{value}{"\""u8}");
    }

    private readonly bool TryWriteReference()
    {
        return buffer.TryWrite(
            $"{" sqref=\""u8}" +
            $"{reference.Reference}" +
            $"{"\">"u8}");
    }

    private bool TryWriteValue1()
    {
        return buffer.TryWrite(
            _progress, out _progress,
            $"{"<formula1>"u8}{validation.Value1}{"</formula1>"u8}");
    }

    private bool TryWriteValue2()
    {
        if (validation.Value2 is null)
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{"<formula2>"u8}{validation.Value2}{"</formula2>"u8}");
    }

    private enum Element
    {
        Header,
        InputTitle,
        InputMessage,
        ErrorTitle,
        ErrorMessage,
        Reference,
        Value1,
        Value2,
        Footer,
        Done
    }
}
