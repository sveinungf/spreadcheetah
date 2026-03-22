using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Validations;

namespace SpreadCheetah.MetadataXml;

internal struct DataValidationXml(
    SingleCellOrCellRangeReference reference,
    DataValidation validation,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;

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
            Element.InputTitleStart => TryWriteInputTitleStart(),
            Element.InputTitle => TryWriteValue(valid.InputTitle),
            Element.InputTitleEnd => TryWriteAttributeEnd(valid.InputTitle),
            Element.InputMessageStart => TryWriteInputMessageStart(),
            Element.InputMessage => TryWriteValue(valid.InputMessage),
            Element.InputMessageEnd => TryWriteAttributeEnd(valid.InputMessage),
            Element.ErrorTitleStart => TryWriteErrorTitleStart(),
            Element.ErrorTitle => TryWriteValue(valid.ErrorTitle),
            Element.ErrorTitleEnd => TryWriteAttributeEnd(valid.ErrorTitle),
            Element.ErrorMessageStart => TryWriteErrorMessageStart(),
            Element.ErrorMessage => TryWriteValue(valid.ErrorMessage),
            Element.ErrorMessageEnd => TryWriteAttributeEnd(valid.ErrorMessage),
            Element.Reference => TryWriteReference(),
            Element.Value1Start => TryWriteValue1Start(),
            Element.Value1 => TryWriteValue(valid.Value1),
            Element.Value1End => TryWriteValue1End(),
            Element.Value2Start => TryWriteValue2Start(),
            Element.Value2 => TryWriteValue(valid.Value2),
            Element.Value2End => TryWriteValue2End(),
            _ => buffer.TryWrite("</dataValidation>"u8)
        };

        if (Current)
            ++_next;

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

    private readonly bool TryWriteInputTitleStart()
        => string.IsNullOrEmpty(validation.InputTitle) || buffer.TryWrite(" promptTitle=\""u8);

    private readonly bool TryWriteInputMessageStart()
        => string.IsNullOrEmpty(validation.InputMessage) || buffer.TryWrite(" prompt=\""u8);

    private readonly bool TryWriteErrorTitleStart()
        => string.IsNullOrEmpty(validation.ErrorTitle) || buffer.TryWrite(" errorTitle=\""u8);

    private readonly bool TryWriteErrorMessageStart()
        => string.IsNullOrEmpty(validation.ErrorMessage) || buffer.TryWrite(" error=\""u8);

    private readonly bool TryWriteAttributeEnd(string? value)
        => string.IsNullOrEmpty(value) || buffer.TryWrite("\""u8);

    private readonly bool TryWriteReference()
    {
        return buffer.TryWrite(
            $"{" sqref=\""u8}" +
            $"{reference.Reference}" +
            $"{"\""u8}");
    }

    private readonly bool TryWriteValue1Start() => buffer.TryWrite("><formula1>"u8);

    private readonly bool TryWriteValue1End() => buffer.TryWrite("</formula1>"u8);

    private readonly bool TryWriteValue2Start()
        => validation.Value2 is null || buffer.TryWrite("<formula2>"u8);

    private readonly bool TryWriteValue2End()
        => validation.Value2 is null || buffer.TryWrite("</formula2>"u8);

    private bool TryWriteValue(string? value)
    {
        if (string.IsNullOrEmpty(value)) return true;
        if (!buffer.WriteLongString(value, ref _nextIndex))
            return false;

        _nextIndex = 0;
        return true;
    }

    private enum Element
    {
        Header,
        InputTitleStart,
        InputTitle,
        InputTitleEnd,
        InputMessageStart,
        InputMessage,
        InputMessageEnd,
        ErrorTitleStart,
        ErrorTitle,
        ErrorTitleEnd,
        ErrorMessageStart,
        ErrorMessage,
        ErrorMessageEnd,
        Reference,
        Value1Start,
        Value1,
        Value1End,
        Value2Start,
        Value2,
        Value2End,
        Footer,
        Done
    }
}