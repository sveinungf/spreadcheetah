using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO;
using Xunit;

namespace SpreadCheetah.Benchmark.Test.Helpers
{
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
    }
}
