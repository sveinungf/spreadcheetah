using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Helpers;

internal static class ThrowHelper
{
    [DoesNotReturn]
    public static void CantAddImageEmbeddedInOtherSpreadsheet() => throw new SpreadCheetahException("The image can't be added because it was embedded in another spreadsheet.");

    [DoesNotReturn]
    public static void ColumnNameInvalid(string? paramName) => throw new ArgumentException("Invalid column name.", paramName);

    [DoesNotReturn]
    public static void ColumnNumberInvalid(string? paramName, int number) => throw new ArgumentOutOfRangeException(paramName, number, "The column number must be greater than 0 and can't be larger than 16384.");

    [DoesNotReturn]
    public static void SingleCellReferenceInvalid(string? paramName) => throw new ArgumentException("Invalid reference for a cell.", paramName);

    [DoesNotReturn]
    public static void CellRangeReferenceInvalid(string? paramName) => throw new ArgumentException("Invalid reference for a cell range.", paramName);

    [DoesNotReturn]
    public static void SingleCellOrCellRangeReferenceInvalid(string? paramName) => throw new ArgumentException("Invalid reference for a cell or a range of cells.", paramName);

    [DoesNotReturn]
    public static void EnumValueInvalid<T>(string? paramName, T value) => throw new ArgumentOutOfRangeException(paramName, value, "The value is not a valid enum value.");

    [DoesNotReturn]
    public static void MaxNumberOfDataValidations() => throw new SpreadCheetahException($"Can't add more than {SpreadsheetConstants.MaxNumberOfDataValidations} data validations to a worksheet.");

    [DoesNotReturn]
    public static void NoActiveWorksheet() => throw new SpreadCheetahException("There is no active worksheet.");

    [DoesNotReturn]
    public static void EmbedImageBeforeStartingWorksheet() => throw new SpreadCheetahException("Images must be embedded before starting a worksheet.");

    [DoesNotReturn]
    public static void EmbedImageNotAllowedAfterFinish() => throw new SpreadCheetahException($"Can't embed image after {nameof(Spreadsheet.FinishAsync)} has been called.");

    [DoesNotReturn]
    public static void FillCellRangeMustContainAtLeastOneCell(string? paramName) => throw new ArgumentOutOfRangeException(paramName, "The cell range must contain at least one cell.");

    [DoesNotReturn]
    public static void ImageDimensionTooLarge(string? paramName, int actualValue) => throw new ArgumentOutOfRangeException(paramName, actualValue, $"Image width or height can't exceed {SpreadsheetConstants.MaxImageDimension} pixels.");

    [DoesNotReturn]
    public static void ImageDimensionZeroOrNegative(string? paramName, int actualValue) => throw new ArgumentOutOfRangeException(paramName, actualValue, "Image width and height must be at least 1 pixel.");

    [DoesNotReturn]
    public static void ImageScaleTooLarge(string? paramName, float actualValue) => throw new ArgumentOutOfRangeException(paramName, actualValue, $"The image scale can't cause image width or height to exceed {SpreadsheetConstants.MaxImageDimension} pixels.");

    [DoesNotReturn]
    public static void ImageScaleTooSmall(string? paramName, float actualValue) => throw new ArgumentOutOfRangeException(paramName, actualValue, "The image scale must result in image width and height being at least 1 pixel.");

    [DoesNotReturn]
    public static void MinGreaterThanMax(string? minParamName) => throw new ArgumentException("The min value must be less than or equal to the max value.", minParamName);

    [DoesNotReturn]
    public static void NameEmptyOrWhiteSpace(string? paramName) => throw new ArgumentException("The name can not be empty or consist only of whitespace.", paramName);

    [DoesNotReturn]
    public static void NameTooLong(int maxLength, string? paramName) => throw new ArgumentException(StringHelper.Invariant($"The name can not be more than {maxLength} characters."), paramName);

    [DoesNotReturn]
    public static void NoteTextTooLong(string? paramName) => throw new ArgumentException($"Note text can not exceed {SpreadsheetConstants.MaxNoteTextLength} characters.", paramName);

    [DoesNotReturn]
    public static void ResizeAndMoveCellsCombinationNotSupported(string resizeParamName, string moveParamName) => throw new ArgumentException($"Enabling {resizeParamName} is not supported when {moveParamName} is false.", resizeParamName);

    [DoesNotReturn]
    public static void SpreadsheetMustContainWorksheet() => throw new SpreadCheetahException("Spreadsheet must contain at least one worksheet.");

    [DoesNotReturn]
    public static void StartWorksheetNotAllowedAfterFinish() => throw new SpreadCheetahException($"Can't start another worksheet after {nameof(Spreadsheet.FinishAsync)} has been called.");

    [DoesNotReturn]
    public static void StreamDoesNotSupportReading(string? paramName) => throw new ArgumentException("The stream does not support reading.", paramName);

    [DoesNotReturn]
    public static void StreamReadNoBytes(string? paramName) => throw new ArgumentException("Could not read any data from the stream. Perhaps the stream's Position should be set to 0 before calling the method?", paramName);

    [DoesNotReturn]
    public static void StreamReadNotEnoughBytes(string? paramName) => throw new ArgumentException("The stream did not contain enough data to determine the actual content.", paramName);

    [DoesNotReturn]
    public static void StreamContentNotSupportedImageType(string? paramName) => throw new ArgumentException("The stream content is not a supported image type. Currently only PNG images are supported.", nameof(paramName));

    [DoesNotReturn]
    public static void StyleNameAlreadyExists(string? paramName) => throw new ArgumentException("A style with the given name already exists.", paramName);

    [DoesNotReturn]
    public static void StyleNameCanNotEqualNormal(string? paramName) => throw new ArgumentException("The name can not be equal to 'Normal', since that is the name of the default built-in style.", paramName);

    [DoesNotReturn]
    public static void StyleNameNotFound(string name) => throw new SpreadCheetahException($"Style with name '{name}' was not found. Make sure the style is first added to the spreadsheet by calling {nameof(Spreadsheet)}.{nameof(Spreadsheet.AddStyle)} with a name argument.");

    [DoesNotReturn]
    public static void StyleNameStartsOrEndsWithWhiteSpace(string? paramName) => throw new ArgumentException("The name can not start or end with white-space.", paramName);

    [DoesNotReturn]
    public static void StyleNameTooLong(string? paramName) => throw new ArgumentException("The name can not be more than 255 characters.", paramName);

    [DoesNotReturn]
    public static void ValueIsNegative<T>(string? paramName, T value) => throw new ArgumentOutOfRangeException(paramName, value, "The value can not be negative.");

    [DoesNotReturn]
    public static void WorksheetNameAlreadyExists(string? paramName) => throw new ArgumentException("A worksheet with the given name already exists.", paramName);

    [DoesNotReturn]
    public static void WorksheetNameInvalidCharacters(string? paramName, string invalidChars) => throw new ArgumentException("The name can not contain any of the following characters: " + invalidChars, paramName);

    [DoesNotReturn]
    public static void WorksheetNameStartsOrEndsWithSingleQuote(string? paramName) => throw new ArgumentException("The name can not start or end with a single quote.", paramName);
}
