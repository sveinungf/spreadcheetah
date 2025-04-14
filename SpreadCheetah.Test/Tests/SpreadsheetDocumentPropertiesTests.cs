using SpreadCheetah.Metadata;
using SpreadCheetah.TestHelpers.Assertions;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetDocumentPropertiesTests
{
    [Fact]
    public async Task Spreadsheet_DocumentProperties_AllPropertiesSet()
    {
        // Arrange
        const string author = "Some author";
        const string subject = "Some subject";
        const string title = "Some title";

        var options = new SpreadCheetahOptions
        {
            BufferSize = SpreadCheetahOptions.MinimumBufferSize,
            DocumentProperties = new DocumentProperties
            {
                Author = author,
                Subject = subject,
                Title = title
            }
        };

        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");
        await spreadsheet.FinishAsync();

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Equal(author, actual.Author);
        Assert.Equal(subject, actual.Subject);
        Assert.Equal(title, actual.Title);
        var actualCreated = Assert.NotNull(actual.Created);
        Assert.Equal(DateTime.UtcNow, actualCreated, TimeSpan.FromSeconds(5));
        Assert.Equal("SpreadCheetah", actual.Application);
    }

    [Fact]
    public async Task Spreadsheet_DocumentProperties_ExplicitlySetToNull()
    {
        // Arrange
        var options = new SpreadCheetahOptions
        {
            DocumentProperties = null
        };

        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");
        await spreadsheet.FinishAsync();

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Null(actual.Author);
        Assert.Null(actual.Subject);
        Assert.Null(actual.Title);
        Assert.Null(actual.Created);
        Assert.Null(actual.Application);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_DocumentProperties_Default(bool withOptions)
    {
        // Arrange
        SpreadCheetahOptions? options = withOptions ? new() : null;
        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");
        await spreadsheet.FinishAsync();

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Null(actual.Author);
        Assert.Null(actual.Subject);
        Assert.Null(actual.Title);
        var actualCreated = Assert.NotNull(actual.Created);
        Assert.Equal(DateTime.UtcNow, actualCreated, TimeSpan.FromSeconds(5));
        Assert.Equal("SpreadCheetah", actual.Application);
    }
}
