using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
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
            Element.InputTitle => TryWriteAttributeValue(valid.InputTitle),
            Element.InputTitleEnd => TryWriteAttributeEnd(valid.InputTitle),
            Element.InputMessageStart => TryWriteInputMessageStart(),
            Element.InputMessage => TryWriteAttributeValue(valid.InputMessage),
            Element.InputMessageEnd => TryWriteAttributeEnd(valid.InputMessage),
            Element.ErrorTitleStart => TryWriteErrorTitleStart(),
            Element.ErrorTitle => TryWriteAttributeValue(valid.ErrorTitle),
            Element.ErrorTitleEnd => TryWriteAttributeEnd(valid.ErrorTitle),
            Element.ErrorMessageStart => TryWriteErrorMessageStart(),
            Element.ErrorMessage => TryWriteAttributeValue(valid.ErrorMessage),
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
        var span = buffer.GetSpan();
        var written = 0;
        var valid = validation;

        if (!"<dataValidation "u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteType(valid.Type, span, ref written)) return false;
        if (!TryWriteErrorType(valid.ErrorType, span, ref written)) return false;
        if (!TryWriteOperator(valid.Operator, span, ref written)) return false;

        if (valid.IgnoreBlank && !"allowBlank=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (!valid.ShowDropdown && !"showDropDown=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (valid.ShowInputMessage && !"showInputMessage=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (valid.ShowErrorAlert && !"showErrorMessage=\"1\" "u8.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private static bool TryWriteType(ValidationType type, Span<byte> bytes, ref int bytesWritten) => type switch
    {
        ValidationType.DateTime => "type=\"date\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.Decimal => "type=\"decimal\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.Integer => "type=\"whole\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.List => "type=\"list\" "u8.TryCopyTo(bytes, ref bytesWritten),
        _ => "type=\"textLength\" "u8.TryCopyTo(bytes, ref bytesWritten)
    };

    private static bool TryWriteErrorType(ValidationErrorType type, Span<byte> bytes, ref int bytesWritten) => type switch
    {
        ValidationErrorType.Warning => "errorStyle=\"warning\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationErrorType.Information => "errorStyle=\"information\" "u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };

    private static bool TryWriteOperator(ValidationOperator op, Span<byte> bytes, ref int bytesWritten) => op switch
    {
        ValidationOperator.NotBetween => "operator=\"notBetween\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.EqualTo => "operator=\"equal\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.NotEqualTo => "operator=\"notEqual\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.GreaterThan => "operator=\"greaterThan\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.LessThan => "operator=\"lessThan\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.GreaterThanOrEqualTo => "operator=\"greaterThanOrEqual\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationOperator.LessThanOrEqualTo => "operator=\"lessThanOrEqual\" "u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };

    private readonly bool TryWriteInputTitleStart()
        => string.IsNullOrEmpty(validation.InputTitle) || buffer.TryWrite("promptTitle=\""u8);

    private readonly bool TryWriteInputMessageStart()
        => string.IsNullOrEmpty(validation.InputMessage) || buffer.TryWrite("prompt=\""u8);

    private readonly bool TryWriteErrorTitleStart()
        => string.IsNullOrEmpty(validation.ErrorTitle) || buffer.TryWrite("errorTitle=\""u8);

    private readonly bool TryWriteErrorMessageStart()
        => string.IsNullOrEmpty(validation.ErrorMessage) || buffer.TryWrite("error=\""u8);

#pragma warning disable EPS12 // A struct member can be made readonly
    private bool TryWriteAttributeValue(string? value)
#pragma warning restore EPS12 // A struct member can be made readonly
    {
        if (string.IsNullOrEmpty(value)) return true;
        var encodedValue = XmlUtility.XmlEncode(value);
        return TryWriteValue(encodedValue);
    }

    private readonly bool TryWriteAttributeEnd(string? value)
        => string.IsNullOrEmpty(value) || buffer.TryWrite("\" "u8);

    private readonly bool TryWriteReference()
    {
        return buffer.TryWrite(
            $"{"sqref=\""u8}" +
            $"{reference.Reference}" +
            $"{"\" "u8}");
    }

    private readonly bool TryWriteValue1Start() => buffer.TryWrite("><formula1>"u8);

    private readonly bool TryWriteValue1End() => buffer.TryWrite("</formula1>"u8);

    private readonly bool TryWriteValue2Start()
        => validation.Value2 is null || buffer.TryWrite("<formula2>"u8);

    private readonly bool TryWriteValue2End()
        => validation.Value2 is null || buffer.TryWrite("</formula2>"u8);

    private bool TryWriteValue(string? value)
    {
        if (value is null) return true;
        var bytes = buffer.GetSpan();
        var bytesWritten = 0;
        if (!SpanHelper.TryWriteLongString(value, ref _nextIndex, bytes, ref bytesWritten))
        {
            buffer.Advance(bytesWritten);
            return false;
        }

        buffer.Advance(bytesWritten);
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