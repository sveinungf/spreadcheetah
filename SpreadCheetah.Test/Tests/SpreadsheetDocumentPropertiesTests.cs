using SpreadCheetah.Metadata;
using SpreadCheetah.TestHelpers.Assertions;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetDocumentPropertiesTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task Spreadsheet_DocumentProperties_AllPropertiesSet()
    {
        // Arrange
        const string author = "Some author";
        const string subject = "Some subject";
        const string title = "Some title";

        var options = new SpreadCheetahOptions
        {
            DocumentProperties = new DocumentProperties
            {
                Author = author,
                Subject = subject,
                Title = title
            }
        };

        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var timestampBeforeFinishing = DateTime.UtcNow;
        await spreadsheet.FinishAsync(Token);

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Equal(author, actual.Author);
        Assert.Equal(subject, actual.Subject);
        Assert.Equal(title, actual.Title);
        var actualCreated = Assert.NotNull(actual.Created);
        Assert.Equal(timestampBeforeFinishing, actualCreated, TimeSpan.FromSeconds(10));
        Assert.Equal("SpreadCheetah", actual.Application);
    }

    [Fact]
    public async Task Spreadsheet_DocumentProperties_VeryLongPropertyValues()
    {
        // Arrange
        var author = string.Join(", ", Enumerable.Repeat("Author", 500));
        var subject = string.Join(", ", Enumerable.Repeat("Subject", 500));
        var title = string.Join(", ", Enumerable.Repeat("Title", 500));

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Equal(author, actual.Author);
        Assert.Equal(subject, actual.Subject);
        Assert.Equal(title, actual.Title);
    }

    [Fact]
    public async Task Spreadsheet_DocumentProperties_PropertiesWithCharactersToEscape()
    {
        // Arrange
        const string author = "Lars & Lenny";
        const string subject = "Date & Time";
        const string title = "Then & Now";

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Equal(author, actual.Author);
        Assert.Equal(subject, actual.Subject);
        Assert.Equal(title, actual.Title);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        await spreadsheet.FinishAsync(Token);

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var timestampBeforeFinishing = DateTime.UtcNow;
        await spreadsheet.FinishAsync(Token);

        // Assert
        var actual = SpreadsheetAssert.DocumentProperties(stream);
        Assert.Null(actual.Author);
        Assert.Null(actual.Subject);
        Assert.Null(actual.Title);
        var actualCreated = Assert.NotNull(actual.Created);
        Assert.Equal(timestampBeforeFinishing, actualCreated, TimeSpan.FromSeconds(10));
        Assert.Equal("SpreadCheetah", actual.Application);
    }
}
