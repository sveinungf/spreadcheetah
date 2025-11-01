using SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;
using SpreadCheetah.TestHelpers.Assertions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using SpreadCheetah.TestHelpers.Extensions;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class InferColumnHeadersTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task InferColumnHeaders_ClassWithMultipleProperties()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            "The ID",
            nameof(ClassWithMultipleProperties.Name),
            "The price"
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }
}
