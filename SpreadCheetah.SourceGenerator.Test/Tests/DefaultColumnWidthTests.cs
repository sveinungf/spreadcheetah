using SpreadCheetah.SourceGenerator.Test.Models.DefaultColumnWidth;
using SpreadCheetah.TestHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class DefaultColumnWidthTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task DefaultColumnWidth_ClassWithDefaultColumnWidth()
    {
        // Arrange
        var typeInfo = DefaultColumnWidthContext.Default.ClassWithDefaultColumnWidth;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        var obj = new ClassWithDefaultColumnWidth { Id = 1, Name = "Test" };

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet", typeInfo, token: Token);
        await spreadsheet.AddAsRowAsync(obj, typeInfo, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(10, sheet.Column("A").Width, 0.005);
        Assert.Equal(20, sheet.Column("B").Width, 0.005);
    }
}
