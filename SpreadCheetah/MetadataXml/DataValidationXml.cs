using SpreadCheetah.Helpers;
using SpreadCheetah.Validations;
using System.Net;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class DataValidationXml
{
    public static async ValueTask WriteAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<CellReference, DataValidation> validations,
        CancellationToken token)
    {
        var sb = new StringBuilder("<dataValidations count=\"");
        sb.Append(validations.Count);
        sb.Append("\">");

        foreach (var keyValue in validations)
        {
            var validation = keyValue.Value;
            sb.Append("<dataValidation ");
            sb.AppendType(validation.Type);
            sb.AppendErrorType(validation.ErrorType);
            sb.AppendOperator(validation.Operator);

            if (validation.IgnoreBlank)
                sb.Append("allowBlank=\"1\" ");

            if (!validation.ShowDropdown)
                sb.Append("showDropDown=\"1\" ");

            if (validation.ShowInputMessage)
                sb.Append("showInputMessage=\"1\" ");

            if (validation.ShowErrorAlert)
                sb.Append("showErrorMessage=\"1\" ");

            if (!string.IsNullOrEmpty(validation.InputTitle))
                sb.AppendTextAttribute("promptTitle", validation.InputTitle!);

            if (!string.IsNullOrEmpty(validation.InputMessage))
                sb.AppendTextAttribute("prompt", validation.InputMessage!);

            if (!string.IsNullOrEmpty(validation.ErrorTitle))
                sb.AppendTextAttribute("errorTitle", validation.ErrorTitle!);

            if (!string.IsNullOrEmpty(validation.ErrorMessage))
                sb.AppendTextAttribute("error", validation.ErrorMessage!);

            sb.AppendTextAttribute("sqref", keyValue.Key.Reference);

            if (validation.Value1 is null)
            {
                sb.Append("/>");
                continue;
            }

            sb.Append("><formula1>").Append(validation.Value1).Append("</formula1>");

            if (validation.Value2 is not null)
                sb.Append("<formula2>").Append(validation.Value2).Append("</formula2>");

            sb.Append("</dataValidation>");
        }

        sb.Append("</dataValidations>");
        await buffer.WriteStringAsync(sb, stream, token).ConfigureAwait(false);
    }

    private static void AppendTextAttribute(this StringBuilder sb, string attribute, string value)
    {
        sb.Append(attribute);
        sb.Append("=\"");
        sb.Append(WebUtility.HtmlEncode(value));
        sb.Append("\" ");
    }

    private static StringBuilder AppendType(this StringBuilder sb, ValidationType type) => type switch
    {
        ValidationType.Decimal => sb.Append("type=\"decimal\" "),
        ValidationType.Integer => sb.Append("type=\"whole\" "),
        ValidationType.List => sb.Append("type=\"list\" "),
        ValidationType.TextLength => sb.Append("type=\"textLength\" "),
        _ => sb
    };

    private static StringBuilder AppendErrorType(this StringBuilder sb, ValidationErrorType type) => type switch
    {
        ValidationErrorType.Warning => sb.Append("errorStyle=\"warning\" "),
        ValidationErrorType.Information => sb.Append("errorStyle=\"information\" "),
        _ => sb
    };

    private static StringBuilder AppendOperator(this StringBuilder sb, ValidationOperator op) => op switch
    {
        ValidationOperator.NotBetween => sb.Append("operator=\"notBetween\" "),
        ValidationOperator.EqualTo => sb.Append("operator=\"equal\" "),
        ValidationOperator.NotEqualTo => sb.Append("operator=\"notEqual\" "),
        ValidationOperator.GreaterThan => sb.Append("operator=\"greaterThan\" "),
        ValidationOperator.LessThan => sb.Append("operator=\"lessThan\" "),
        ValidationOperator.GreaterThanOrEqualTo => sb.Append("operator=\"greaterThanOrEqual\" "),
        ValidationOperator.LessThanOrEqualTo => sb.Append("operator=\"lessThanOrEqual\" "),
        _ => sb
    };
}

internal struct DataValidationXml2
{
    private readonly CellReference _reference;
    private readonly DataValidation _validation;
    private Element _next;
    private int _nextIndex;

    public DataValidationXml2(CellReference reference, DataValidation validation)
    {
        _reference = reference;
        _validation = validation;
    }

    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
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