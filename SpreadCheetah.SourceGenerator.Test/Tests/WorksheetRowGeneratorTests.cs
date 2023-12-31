using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models;
using SpreadCheetah.SourceGenerator.Test.Models.Accessibility;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using SpreadCheetah.SourceGenerator.Test.Models.MultipleProperties;
using SpreadCheetah.SourceGenerator.Test.Models.NoProperties;
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
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(firstName, sheet["A", 1].StringValue);
        Assert.Equal(lastName, sheet["B", 1].StringValue);
        Assert.Equal(age, sheet["C", 1].IntValue);
        Assert.Equal(3, sheet.CellCount);
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
        var cells = sheetPart?.Worksheet.Descendants<OpenXmlCell>();
        Assert.NotNull(cells);
        Assert.Empty(cells);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddAsRow_InternalClassWithSingleProperty(bool explicitInternal)
    {
        // Arrange
        const string name = "Nordmann";

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            if (explicitInternal)
            {
                var obj = new InternalAccessibilityClassWithSingleProperty { Name = name };
                var ctx = InternalAccessibilityContext.Default.InternalAccessibilityClassWithSingleProperty;
                await s.AddAsRowAsync(obj, ctx);
            }
            else
            {
                var obj = new DefaultAccessibilityClassWithSingleProperty { Name = name };
                var ctx = DefaultAccessibilityContext.Default.DefaultAccessibilityClassWithSingleProperty;
                await s.AddAsRowAsync(obj, ctx);
            }

            await s.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var cell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(name, cell.InnerText);
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_ClassWithAllSupportedTypes()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        var obj = new ClassWithAllSupportedTypes();
        var ctx = AllSupportedTypesContext.Default.ClassWithAllSupportedTypes;
        await spreadsheet.AddAsRowAsync(obj, ctx);
        await spreadsheet.FinishAsync();

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var cellCount = sheetPart.Worksheet.Descendants<OpenXmlCell>().Count();
        Assert.Equal(16, cellCount);
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
        var cells = sheetPart?.Worksheet.Descendants<OpenXmlCell>();
        Assert.NotNull(cells);
        Assert.Empty(cells);
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

    [Fact]
    public async Task Spreadsheet_AddAsRow_ThrowsWhenNoWorksheet()
    {
        // Arrange
        var obj = new ClassWithMultipleProperties("Ola", "Nordmann", 25);
        var typeInfo = MultiplePropertiesContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        // Act & Assert
        await Assert.ThrowsAsync<SpreadCheetahException>(() => spreadsheet.AddAsRowAsync(obj, typeInfo).AsTask());
    }

    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.RecordClass)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.RecordStruct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    [InlineData(ObjectType.ReadOnlyRecordStruct)]
    public async Task Spreadsheet_AddRangeAsRows_ObjectWithMultipleProperties(ObjectType type)
    {
        // Arrange
        var values = new (string FirstName, string LastName, int Age)[]
        {
            ("Ola", "Nordmann", 30),
            ("Ingrid", "Hansen", 28),
            ("Oskar", "Berg", 29)
        };

        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync(values.Select(x => new ClassWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ClassWithMultipleProperties),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync(values.Select(x => new RecordClassWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.RecordClassWithMultipleProperties),
                ObjectType.Struct => s.AddRangeAsRowsAsync(values.Select(x => new StructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.StructWithMultipleProperties),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new RecordStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.RecordStructWithMultipleProperties),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ReadOnlyStructWithMultipleProperties),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyRecordStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ReadOnlyRecordStructWithMultipleProperties),
                _ => throw new NotImplementedException()
            };

            await task;
            await s.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var rows = sheetPart.Worksheet.Descendants<Row>().ToList();

        var row1Cells = rows[0].Descendants<OpenXmlCell>().ToList();
        Assert.Equal(values[0].FirstName, row1Cells[0].InnerText);
        Assert.Equal(values[0].LastName, row1Cells[1].InnerText);
        Assert.Equal(values[0].Age.ToString(), row1Cells[2].InnerText);
        Assert.Equal(3, row1Cells.Count);

        var row2Cells = rows[1].Descendants<OpenXmlCell>().ToList();
        Assert.Equal(values[1].FirstName, row2Cells[0].InnerText);
        Assert.Equal(values[1].LastName, row2Cells[1].InnerText);
        Assert.Equal(values[1].Age.ToString(), row2Cells[2].InnerText);
        Assert.Equal(3, row2Cells.Count);

        var row3Cells = rows[2].Descendants<OpenXmlCell>().ToList();
        Assert.Equal(values[2].FirstName, row3Cells[0].InnerText);
        Assert.Equal(values[2].LastName, row3Cells[1].InnerText);
        Assert.Equal(values[2].Age.ToString(), row3Cells[2].InnerText);
        Assert.Equal(3, row3Cells.Count);

        Assert.Equal(3, rows.Count);
    }

    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.RecordClass)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.RecordStruct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    [InlineData(ObjectType.ReadOnlyRecordStruct)]
    public async Task Spreadsheet_AddRangeAsRows_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        var range = Enumerable.Range(0, 3).ToList();
        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync(range.Select(_ => new ClassWithNoProperties()), ctx.ClassWithNoProperties),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync(range.Select(_ => new RecordClassWithNoProperties()), ctx.RecordClassWithNoProperties),
                ObjectType.Struct => s.AddRangeAsRowsAsync(range.Select(_ => new StructWithNoProperties()), ctx.StructWithNoProperties),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync(range.Select(_ => new RecordStructWithNoProperties()), ctx.RecordStructWithNoProperties),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(range.Select(_ => new ReadOnlyStructWithNoProperties()), ctx.ReadOnlyStructWithNoProperties),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(range.Select(_ => new ReadOnlyRecordStructWithNoProperties()), ctx.ReadOnlyRecordStructWithNoProperties),
                _ => throw new NotImplementedException()
            };

            await task;
            await s.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var rows = sheetPart.Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Empty(row.Descendants<OpenXmlCell>()));
    }

    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.RecordClass)]
    [InlineData(ObjectType.Struct)]
    [InlineData(ObjectType.RecordStruct)]
    [InlineData(ObjectType.ReadOnlyStruct)]
    [InlineData(ObjectType.ReadOnlyRecordStruct)]
    public async Task Spreadsheet_AddRangeAsRows_EmptyEnumerable(ObjectType type)
    {
        // Arrange
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream))
        {
            await s.StartWorksheetAsync("Sheet");

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync(Enumerable.Empty<ClassWithMultipleProperties>(), ctx.ClassWithMultipleProperties),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync(Enumerable.Empty<RecordClassWithMultipleProperties>(), ctx.RecordClassWithMultipleProperties),
                ObjectType.Struct => s.AddRangeAsRowsAsync(Enumerable.Empty<StructWithMultipleProperties>(), ctx.StructWithMultipleProperties),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync(Enumerable.Empty<RecordStructWithMultipleProperties>(), ctx.RecordStructWithMultipleProperties),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(Enumerable.Empty<ReadOnlyStructWithMultipleProperties>(), ctx.ReadOnlyStructWithMultipleProperties),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(Enumerable.Empty<ReadOnlyRecordStructWithMultipleProperties>(), ctx.ReadOnlyRecordStructWithMultipleProperties),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync();
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var rows = sheetPart.Worksheet.Descendants<Row>();
        Assert.Empty(rows);
    }
}
