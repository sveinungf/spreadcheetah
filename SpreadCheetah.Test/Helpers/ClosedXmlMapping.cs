using ClosedXML.Excel;
using SpreadCheetah.Validations;

namespace SpreadCheetah.Test.Helpers;

internal static class ClosedXmlMapping
{
    public static XLErrorStyle GetClosedXmlErrorStyle(this ValidationErrorType type)
    {
        return type switch
        {
            ValidationErrorType.Blocking => XLErrorStyle.Stop,
            ValidationErrorType.Warning => XLErrorStyle.Warning,
            ValidationErrorType.Information => XLErrorStyle.Information,
            _ => throw new NotImplementedException(),
        };
    }
}
