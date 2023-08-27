using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal struct DataValidationXml
{
    private readonly SingleCellOrCellRangeReference _reference;
    private readonly DataValidation _validation;
    private Element _next;
    private int _nextIndex;

    public DataValidationXml(SingleCellOrCellRangeReference reference, DataValidation validation)
    {
        _reference = reference;
        _validation = validation;
    }

#pragma warning disable MA0051 // Method is too long
    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
#pragma warning restore MA0051 // Method is too long
    {
        var valid = _validation;

        if (_next == Element.Header && !Advance(TryWriteHeader(bytes, ref bytesWritten))) return false;
        if (_next == Element.InputTitleStart && !Advance(TryWriteInputTitleStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.InputTitle && !Advance(TryWriteAttributeValue(valid.InputTitle, bytes, ref bytesWritten))) return false;
        if (_next == Element.InputTitleEnd && !Advance(TryWriteAttributeEnd(valid.InputTitle, bytes, ref bytesWritten))) return false;
        if (_next == Element.InputMessageStart && !Advance(TryWriteInputMessageStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.InputMessage && !Advance(TryWriteAttributeValue(valid.InputMessage, bytes, ref bytesWritten))) return false;
        if (_next == Element.InputMessageEnd && !Advance(TryWriteAttributeEnd(valid.InputMessage, bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorTitleStart && !Advance(TryWriteErrorTitleStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorTitle && !Advance(TryWriteAttributeValue(valid.ErrorTitle, bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorTitleEnd && !Advance(TryWriteAttributeEnd(valid.ErrorTitle, bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorMessageStart && !Advance(TryWriteErrorMessageStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorMessage && !Advance(TryWriteAttributeValue(valid.ErrorMessage, bytes, ref bytesWritten))) return false;
        if (_next == Element.ErrorMessageEnd && !Advance(TryWriteAttributeEnd(valid.ErrorMessage, bytes, ref bytesWritten))) return false;
        if (_next == Element.Reference && !Advance(TryWriteReference(bytes, ref bytesWritten))) return false;
        if (_next == Element.Value1Start && !Advance(TryWriteValue1Start(bytes, ref bytesWritten))) return false;
        if (_next == Element.Value1 && !Advance(TryWriteValue(valid.Value1, bytes, ref bytesWritten))) return false;
        if (_next == Element.Value1End && !Advance(TryWriteValue1End(bytes, ref bytesWritten))) return false;
        if (_next == Element.Value2Start && !Advance(TryWriteValue2Start(bytes, ref bytesWritten))) return false;
        if (_next == Element.Value2 && !Advance(TryWriteValue(valid.Value2, bytes, ref bytesWritten))) return false;
        if (_next == Element.Value2End && !Advance(TryWriteValue2End(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance("</dataValidation>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteHeader(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;
        var valid = _validation;

        if (!"<dataValidation "u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteType(valid.Type, span, ref written)) return false;
        if (!TryWriteErrorType(valid.ErrorType, span, ref written)) return false;
        if (!TryWriteOperator(valid.Operator, span, ref written)) return false;

        if (valid.IgnoreBlank && !"allowBlank=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (!valid.ShowDropdown && !"showDropDown=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (valid.ShowInputMessage && !"showInputMessage=\"1\" "u8.TryCopyTo(span, ref written)) return false;
        if (valid.ShowErrorAlert && !"showErrorMessage=\"1\" "u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private static bool TryWriteType(ValidationType type, Span<byte> bytes, ref int bytesWritten) => type switch
    {
        ValidationType.Decimal => "type=\"decimal\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.Integer => "type=\"whole\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.List => "type=\"list\" "u8.TryCopyTo(bytes, ref bytesWritten),
        ValidationType.TextLength => "type=\"textLength\" "u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
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

    private readonly bool TryWriteInputTitleStart(Span<byte> bytes, ref int bytesWritten)
        => string.IsNullOrEmpty(_validation.InputTitle) || "promptTitle=\""u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteInputMessageStart(Span<byte> bytes, ref int bytesWritten)
        => string.IsNullOrEmpty(_validation.InputMessage) || "prompt=\""u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteErrorTitleStart(Span<byte> bytes, ref int bytesWritten)
        => string.IsNullOrEmpty(_validation.ErrorTitle) || "errorTitle=\""u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteErrorMessageStart(Span<byte> bytes, ref int bytesWritten)
        => string.IsNullOrEmpty(_validation.ErrorMessage) || "error=\""u8.TryCopyTo(bytes, ref bytesWritten);

    private bool TryWriteAttributeValue(string? value, Span<byte> bytes, ref int bytesWritten)
    {
        if (string.IsNullOrEmpty(value)) return true;

        var encodedValue = WebUtility.HtmlEncode(value);
        if (!SpanHelper.TryWriteLongString(encodedValue, ref _nextIndex, bytes, ref bytesWritten)) return false;

        _nextIndex = 0;
        return true;
    }

    private static bool TryWriteAttributeEnd(string? value, Span<byte> bytes, ref int bytesWritten)
        => string.IsNullOrEmpty(value) || "\" "u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteReference(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"sqref=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_reference.Reference, span, ref written)) return false;
        if (!"\" "u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private readonly bool TryWriteValue1Start(Span<byte> bytes, ref int bytesWritten)
        => _validation.Value1 is null
            ? "/>"u8.TryCopyTo(bytes, ref bytesWritten)
            : "><formula1>"u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteValue1End(Span<byte> bytes, ref int bytesWritten)
        => _validation.Value1 is null || "</formula1>"u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteValue2Start(Span<byte> bytes, ref int bytesWritten)
        => _validation.Value2 is null || "<formula2>"u8.TryCopyTo(bytes, ref bytesWritten);

    private readonly bool TryWriteValue2End(Span<byte> bytes, ref int bytesWritten)
        => _validation.Value2 is null || "</formula2>"u8.TryCopyTo(bytes, ref bytesWritten);

    private bool TryWriteValue(string? value, Span<byte> bytes, ref int bytesWritten)
    {
        if (value is null) return true;
        if (!SpanHelper.TryWriteLongString(value, ref _nextIndex, bytes, ref bytesWritten)) return false;

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