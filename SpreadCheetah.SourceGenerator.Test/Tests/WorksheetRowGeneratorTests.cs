using DocumentFormat.OpenXml.Packaging;
using SpreadCheetah.SourceGenerator.Test.Models;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using Xunit;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class WorksheetRowGeneratorTests
{
    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.Record)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    public async Task Spreadsheet_AddAsRow_ObjectWithMultipleProperties(ObjectType type)
    {
        // Arrange
        const string firstName = "Ola";
        const string lastName = "Nordmann";
        const int age = 30;
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            if (type == ObjectType.Class)
                await spreadsheet.AddAsRowAsync(new ClassWithMultipleProperties(firstName, lastName, age), ctx.ClassWithMultipleProperties);
            else if (type == ObjectType.Record)
                await spreadsheet.AddAsRowAsync(new RecordClassWithMultipleProperties(firstName, lastName, age), ctx.RecordClassWithMultipleProperties);
            else if (type == ObjectType.Struct)
                await spreadsheet.AddAsRowAsync(new StructWithMultipleProperties(firstName, lastName, age), ctx.StructWithMultipleProperties);
            else if (type == ObjectType.ReadOnlyStruct)
                await spreadsheet.AddAsRowAsync(new ReadOnlyStructWithMultipleProperties(firstName, lastName, age), ctx.ReadOnlyStructWithMultipleProperties);

            await spreadsheet.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var cells = sheetPart.Worksheet.Descendants<OpenXmlCell>().ToList();
        Assert.Equal(firstName, cells[0].InnerText);
        Assert.Equal(lastName, cells[1].InnerText);
        Assert.Equal(age.ToString(), cells[2].InnerText);
        Assert.Equal(3, cells.Count);
    }

    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.Record)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    public async Task Spreadsheet_AddAsRow_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            if (type == ObjectType.Class)
                await spreadsheet.AddAsRowAsync(new ClassWithNoProperties(), ctx.ClassWithNoProperties);
            else if (type == ObjectType.Record)
                await spreadsheet.AddAsRowAsync(new RecordClassWithNoProperties(), ctx.RecordClassWithNoProperties);
            else if (type == ObjectType.Struct)
                await spreadsheet.AddAsRowAsync(new StructWithNoProperties(), ctx.StructWithNoProperties);
            else if (type == ObjectType.ReadOnlyStruct)
                await spreadsheet.AddAsRowAsync(new ReadOnlyStructWithNoProperties(), ctx.ReadOnlyStructWithNoProperties);

            await spreadsheet.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart?.WorksheetParts.Single();
        Assert.Empty(sheetPart?.Worksheet.Descendants<OpenXmlCell>());
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_ObjectWithCustomType()
    {
        // Arrange
        const string value = "value";
        var customType = new CustomType("The name");
        var obj = new RecordClassWithCustomType(customType, value);
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            await spreadsheet.AddAsRowAsync(obj, CustomTypeContext.Default.RecordClassWithCustomType);

            await spreadsheet.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var cells = sheetPart.Worksheet.Descendants<OpenXmlCell>().ToList();
        var actualCell = Assert.Single(cells);
        Assert.Equal(value, actualCell.InnerText);
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_NullObject()
    {
        // Arrange
        ClassWithMultipleProperties obj = null!;
        var typeInfo = MultiplePropertiesContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            await spreadsheet.AddAsRowAsync(obj, typeInfo);

            await spreadsheet.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart?.WorksheetParts.Single();
        Assert.Empty(sheetPart?.Worksheet.Descendants<OpenXmlCell>());
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_NullTypeInfo()
    {
        // Arrange
        var obj = new ClassWithMultipleProperties("Ola", "Nordmann", 25);
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentNullException>(() => spreadsheet.AddAsRowAsync(obj, null!).AsTask());
    }
}
