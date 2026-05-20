using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SpreadCheetah.Validations;

namespace SpreadCheetah.Test.Helpers;

internal static class SpreadsheetAssert
{
    public static void Valid(Stream stream)
    {
        stream.Position = 0;

        using var document = SpreadsheetDocument.Open(stream, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document);
        Assert.Empty(errors);

        stream.Position = 0;
    }

    public static void EquivalentDataValidation(DataValidation validation, IXLDataValidation closedXmlValidation)
    {
        Assert.Equal(validation.ErrorMessage ?? "", closedXmlValidation.ErrorMessage);
        Assert.Equal(validation.ErrorTitle ?? "", closedXmlValidation.ErrorTitle);
        Assert.Equal(validation.ErrorType.GetClosedXmlErrorStyle(), closedXmlValidation.ErrorStyle);
        Assert.Equal(validation.InputMessage ?? "", closedXmlValidation.InputMessage);
        Assert.Equal(validation.InputTitle ?? "", closedXmlValidation.InputTitle);

        // ClosedXml seems to ignore the boolean attributes when reading, and they always become true.
        // If this were to change, then these should be compared against the equivalent on DataValidation.
        Assert.True(closedXmlValidation.IgnoreBlanks);
        Assert.True(closedXmlValidation.ShowErrorMessage);
        Assert.True(closedXmlValidation.ShowInputMessage);
    }
}
