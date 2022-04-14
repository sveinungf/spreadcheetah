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
    [InlineData(ObjectType.RecordClass)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.RecordStruct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    [InlineData(ObjectType.ReadOnlyRecordStruct)]
    public async Task Spreadsheet_AddAsRow_ObjectWithMultipleProperties(ObjectType type)
    {
        // Arrange
        const string firstName = "Ola";
        const string lastName = "Nordmann";
        const int age = 30;
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddAsRowAsync(new ClassWithMultipleProperties(firstName, lastName, age), ctx.ClassWithMultipleProperties),
                ObjectType.RecordClass => s.AddAsRowAsync(new RecordClassWithMultipleProperties(firstName, lastName, age), ctx.RecordClassWithMultipleProperties),
                ObjectType.Struct => s.AddAsRowAsync(new StructWithMultipleProperties(firstName, lastName, age), ctx.StructWithMultipleProperties),
                ObjectType.RecordStruct => s.AddAsRowAsync(new RecordStructWithMultipleProperties(firstName, lastName, age), ctx.RecordStructWithMultipleProperties),
                ObjectType.ReadOnlyStruct => s.AddAsRowAsync(new ReadOnlyStructWithMultipleProperties(firstName, lastName, age), ctx.ReadOnlyStructWithMultipleProperties),
                ObjectType.ReadOnlyRecordStruct => s.AddAsRowAsync(new ReadOnlyRecordStructWithMultipleProperties(firstName, lastName, age), ctx.ReadOnlyRecordStructWithMultipleProperties),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync();
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
    [InlineData(ObjectType.RecordClass)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.RecordStruct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    [InlineData(ObjectType.ReadOnlyRecordStruct)]
    public async Task Spreadsheet_AddAsRow_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddAsRowAsync(new ClassWithNoProperties(), ctx.ClassWithNoProperties),
                ObjectType.RecordClass => s.AddAsRowAsync(new RecordClassWithNoProperties(), ctx.RecordClassWithNoProperties),
                ObjectType.Struct => s.AddAsRowAsync(new StructWithNoProperties(), ctx.StructWithNoProperties),
                ObjectType.RecordStruct => s.AddAsRowAsync(new RecordStructWithNoProperties(), ctx.RecordStructWithNoProperties),
                ObjectType.ReadOnlyStruct => s.AddAsRowAsync(new ReadOnlyStructWithNoProperties(), ctx.ReadOnlyStructWithNoProperties),
                ObjectType.ReadOnlyRecordStruct => s.AddAsRowAsync(new ReadOnlyRecordStructWithNoProperties(), ctx.ReadOnlyRecordStructWithNoProperties),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync();
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
