using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models;
using SpreadCheetah.SourceGenerator.Test.Models.Accessibility;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnWidth;
using SpreadCheetah.SourceGenerator.Test.Models.Combinations;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using SpreadCheetah.SourceGenerator.Test.Models.MultipleProperties;
using SpreadCheetah.SourceGenerator.Test.Models.NoProperties;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using SpreadCheetah.TestHelpers.Extensions;
using Xunit;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class WorksheetRowGeneratorTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddAsRow_ObjectWithMultipleProperties(ObjectType type)
    {
        // Arrange
        const string firstName = "Ola";
        const string lastName = "Nordmann";
        const int age = 30;
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddAsRowAsync(new ClassWithMultipleProperties(firstName, lastName, age), ctx.ClassWithMultipleProperties, Token),
                ObjectType.RecordClass => s.AddAsRowAsync(new RecordClassWithMultipleProperties(firstName, lastName, age), ctx.RecordClassWithMultipleProperties, Token),
                ObjectType.Struct => s.AddAsRowAsync(new StructWithMultipleProperties(firstName, lastName, age), ctx.StructWithMultipleProperties, Token),
                ObjectType.RecordStruct => s.AddAsRowAsync(new RecordStructWithMultipleProperties(firstName, lastName, age), ctx.RecordStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyStruct => s.AddAsRowAsync(new ReadOnlyStructWithMultipleProperties(firstName, lastName, age), ctx.ReadOnlyStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyRecordStruct => s.AddAsRowAsync(new ReadOnlyRecordStructWithMultipleProperties(firstName, lastName, age), ctx.ReadOnlyRecordStructWithMultipleProperties, Token),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync(Token);
        }

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(firstName, sheet["A", 1].StringValue);
        Assert.Equal(lastName, sheet["B", 1].StringValue);
        Assert.Equal(age, sheet["C", 1].IntValue);
        Assert.Equal(3, sheet.CellCount);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddAsRow_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddAsRowAsync(new ClassWithNoProperties(), ctx.ClassWithNoProperties, Token),
                ObjectType.RecordClass => s.AddAsRowAsync(new RecordClassWithNoProperties(), ctx.RecordClassWithNoProperties, Token),
                ObjectType.Struct => s.AddAsRowAsync(new StructWithNoProperties(), ctx.StructWithNoProperties, Token),
                ObjectType.RecordStruct => s.AddAsRowAsync(new RecordStructWithNoProperties(), ctx.RecordStructWithNoProperties, Token),
                ObjectType.ReadOnlyStruct => s.AddAsRowAsync(new ReadOnlyStructWithNoProperties(), ctx.ReadOnlyStructWithNoProperties, Token),
                ObjectType.ReadOnlyRecordStruct => s.AddAsRowAsync(new ReadOnlyRecordStructWithNoProperties(), ctx.ReadOnlyRecordStructWithNoProperties, Token),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync(Token);
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
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            if (explicitInternal)
            {
                var obj = new InternalAccessibilityClassWithSingleProperty { Name = name };
                var ctx = InternalAccessibilityContext.Default.InternalAccessibilityClassWithSingleProperty;
                await s.AddAsRowAsync(obj, ctx, Token);
            }
            else
            {
                var obj = new DefaultAccessibilityClassWithSingleProperty { Name = name };
                var ctx = DefaultAccessibilityContext.Default.DefaultAccessibilityClassWithSingleProperty;
                await s.AddAsRowAsync(obj, ctx, Token);
            }

            await s.FinishAsync(Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var obj = new ClassWithAllSupportedTypes();
        var ctx = AllSupportedTypesContext.Default.ClassWithAllSupportedTypes;
        await spreadsheet.AddAsRowAsync(obj, ctx, Token);
        await spreadsheet.FinishAsync(Token);

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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

            // Act
            await spreadsheet.AddAsRowAsync(obj, CustomTypeContext.Default.RecordClassWithCustomType, Token);

            await spreadsheet.FinishAsync(Token);
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

            // Act
            await spreadsheet.AddAsRowAsync(obj, typeInfo, Token);

            await spreadsheet.FinishAsync(Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentNullException>(() => spreadsheet.AddAsRowAsync(obj, null!, Token).AsTask());
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_ThrowsWhenNoWorksheet()
    {
        // Arrange
        var obj = new ClassWithMultipleProperties("Ola", "Nordmann", 25);
        var typeInfo = MultiplePropertiesContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act & Assert
        await Assert.ThrowsAsync<SpreadCheetahException>(() => spreadsheet.AddAsRowAsync(obj, typeInfo, Token).AsTask());
    }

    [Theory, CombinatorialData]
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
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync(values.Select(x => new ClassWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ClassWithMultipleProperties, Token),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync(values.Select(x => new RecordClassWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.RecordClassWithMultipleProperties, Token),
                ObjectType.Struct => s.AddRangeAsRowsAsync(values.Select(x => new StructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.StructWithMultipleProperties, Token),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new RecordStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.RecordStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ReadOnlyStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyRecordStructWithMultipleProperties(x.FirstName, x.LastName, x.Age)), ctx.ReadOnlyRecordStructWithMultipleProperties, Token),
                _ => throw new NotImplementedException()
            };

            await task;
            await s.FinishAsync(Token);
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

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddRangeAsRows_ObjectWithColumnOrdering(ObjectType type)
    {
        // Arrange
        var values = new (string FirstName, string LastName, decimal Gpa, int Age)[]
        {
            ("Ola", "Nordmann", 3.2m, 30),
            ("Ingrid", "Hansen", 3.4m, 28),
            ("Oskar", "Berg", 2.8m, 29)
        };

        var ctx = ColumnOrderingContext.Default;

        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var task = type switch
        {
            ObjectType.Class => s.AddRangeAsRowsAsync(values.Select(x => new ClassWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.ClassWithColumnOrdering, Token),
            ObjectType.RecordClass => s.AddRangeAsRowsAsync(values.Select(x => new RecordClassWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.RecordClassWithColumnOrdering, Token),
            ObjectType.Struct => s.AddRangeAsRowsAsync(values.Select(x => new StructWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.StructWithColumnOrdering, Token),
            ObjectType.RecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new RecordStructWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.RecordStructWithColumnOrdering, Token),
            ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyStructWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.ReadOnlyStructWithColumnOrdering, Token),
            ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(values.Select(x => new ReadOnlyRecordStructWithColumnOrdering(x.FirstName, x.LastName, x.Gpa, x.Age)), ctx.ReadOnlyRecordStructWithColumnOrdering, Token),
            _ => throw new NotImplementedException()
        };

        await task;
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(values.Select(x => x.LastName), sheet.Column("A").Cells.StringValues());
        Assert.Equal(values.Select(x => x.FirstName), sheet.Column("B").Cells.StringValues());
        Assert.Equal(values.Select(x => x.Age), sheet.Column("C").Cells.Select(x => x.IntValue ?? -1));
        Assert.Equal(values.Select(x => x.Gpa), sheet.Column("D").Cells.Select(x => x.DecimalValue ?? -1));
        Assert.Equal(3, sheet.RowCount);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddRangeAsRows_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        var range = Enumerable.Range(0, 3).ToList();
        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync(range.Select(_ => new ClassWithNoProperties()), ctx.ClassWithNoProperties, Token),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync(range.Select(_ => new RecordClassWithNoProperties()), ctx.RecordClassWithNoProperties, Token),
                ObjectType.Struct => s.AddRangeAsRowsAsync(range.Select(_ => new StructWithNoProperties()), ctx.StructWithNoProperties, Token),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync(range.Select(_ => new RecordStructWithNoProperties()), ctx.RecordStructWithNoProperties, Token),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync(range.Select(_ => new ReadOnlyStructWithNoProperties()), ctx.ReadOnlyStructWithNoProperties, Token),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync(range.Select(_ => new ReadOnlyRecordStructWithNoProperties()), ctx.ReadOnlyRecordStructWithNoProperties, Token),
                _ => throw new NotImplementedException()
            };

            await task;
            await s.FinishAsync(Token);
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var rows = sheetPart.Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Empty(row.Descendants<OpenXmlCell>()));
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddRangeAsRows_EmptyEnumerable(ObjectType type)
    {
        // Arrange
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using (var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await s.StartWorksheetAsync("Sheet", token: Token);

            // Act
            var task = type switch
            {
                ObjectType.Class => s.AddRangeAsRowsAsync([], ctx.ClassWithMultipleProperties, Token),
                ObjectType.RecordClass => s.AddRangeAsRowsAsync([], ctx.RecordClassWithMultipleProperties, Token),
                ObjectType.Struct => s.AddRangeAsRowsAsync([], ctx.StructWithMultipleProperties, Token),
                ObjectType.RecordStruct => s.AddRangeAsRowsAsync([], ctx.RecordStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyStruct => s.AddRangeAsRowsAsync([], ctx.ReadOnlyStructWithMultipleProperties, Token),
                ObjectType.ReadOnlyRecordStruct => s.AddRangeAsRowsAsync([], ctx.ReadOnlyRecordStructWithMultipleProperties, Token),
                _ => throw new NotImplementedException(),
            };

            await task;
            await s.FinishAsync(Token);
        }

        // Assert
        stream.Position = 0;
        using var actual = SpreadsheetDocument.Open(stream, false);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var rows = sheetPart.Worksheet.Descendants<Row>();
        Assert.Empty(rows);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_ObjectWithMultipleProperties(ObjectType type)
    {
        // Arrange
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var task = type switch
        {
            ObjectType.Class => s.AddHeaderRowAsync(ctx.ClassWithMultipleProperties, token: Token),
            ObjectType.RecordClass => s.AddHeaderRowAsync(ctx.RecordClassWithMultipleProperties, token: Token),
            ObjectType.Struct => s.AddHeaderRowAsync(ctx.StructWithMultipleProperties, token: Token),
            ObjectType.RecordStruct => s.AddHeaderRowAsync(ctx.RecordStructWithMultipleProperties, token: Token),
            ObjectType.ReadOnlyStruct => s.AddHeaderRowAsync(ctx.ReadOnlyStructWithMultipleProperties, token: Token),
            ObjectType.ReadOnlyRecordStruct => s.AddHeaderRowAsync(ctx.ReadOnlyRecordStructWithMultipleProperties, token: Token),
            _ => throw new NotImplementedException(),
        };

        await task;
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("FirstName", sheet["A", 1].StringValue);
        Assert.Equal("LastName", sheet["B", 1].StringValue);
        Assert.Equal("Age", sheet["C", 1].StringValue);
        Assert.Equal(3, sheet.CellCount);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_ObjectWithNoProperties(ObjectType type)
    {
        // Arrange
        var ctx = NoPropertiesContext.Default;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var task = type switch
        {
            ObjectType.Class => s.AddHeaderRowAsync(ctx.ClassWithNoProperties, token: Token),
            ObjectType.RecordClass => s.AddHeaderRowAsync(ctx.RecordClassWithNoProperties, token: Token),
            ObjectType.Struct => s.AddHeaderRowAsync(ctx.StructWithNoProperties, token: Token),
            ObjectType.RecordStruct => s.AddHeaderRowAsync(ctx.RecordStructWithNoProperties, token: Token),
            ObjectType.ReadOnlyStruct => s.AddHeaderRowAsync(ctx.ReadOnlyStructWithNoProperties, token: Token),
            ObjectType.ReadOnlyRecordStruct => s.AddHeaderRowAsync(ctx.ReadOnlyRecordStructWithNoProperties, token: Token),
            _ => throw new NotImplementedException(),
        };

        await task;
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Empty(sheet.Row(1).Cells);
        Assert.Equal(0, sheet.CellCount);
    }

    [Fact]
    public async Task Spreadsheet_AddHeaderRow_Styling()
    {
        // Arrange
        var ctx = MultiplePropertiesContext.Default;

        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        var style = new Style { Font = { Bold = true } };
        var styleId = s.AddStyle(style);

        // Act
        await s.AddHeaderRowAsync(ctx.ClassWithMultipleProperties, styleId, Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("FirstName", sheet["A", 1].StringValue);
        Assert.Equal("LastName", sheet["B", 1].StringValue);
        Assert.Equal("Age", sheet["C", 1].StringValue);
        Assert.All(sheet.Row(1).Cells, x => Assert.True(x.Style.Font.Bold));
        Assert.Equal(3, sheet.CellCount);
    }

    [Fact]
    public async Task Spreadsheet_AddHeaderRow_NullTypeInfo()
    {
        // Arrange
        WorksheetRowTypeInfo<ClassWithMultipleProperties> typeInfo = null!;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentNullException>(() => spreadsheet.AddHeaderRowAsync(typeInfo, token: Token).AsTask());
    }

    [Fact]
    public async Task Spreadsheet_AddHeaderRow_ThrowsWhenNoWorksheet()
    {
        // Arrange
        var typeInfo = MultiplePropertiesContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act & Assert
        await Assert.ThrowsAsync<SpreadCheetahException>(() => spreadsheet.AddHeaderRowAsync(typeInfo, token: Token).AsTask());
    }

    [Fact]
    public async Task Spreadsheet_AddHeaderRow_ObjectWithMultipleColumnAttributes()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        IList<string> expectedValues =
        [
            "Year",
            "The make",
            "Country of origin",
            "Model",
            "kW",
            "Length (in cm)",
            "Id"
        ];

        // Act
        await s.AddHeaderRowAsync(ColumnAttributesContext.Default.ClassWithColumnAttributes, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_ObjectWithMultipleColumnAttributes()
    {
        // Arrange
        var ctx = ColumnAttributesContext.Default.ClassWithColumnAttributes;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", ctx, Token);
        s.AddStyle(new Style { Font = { Bold = true } }, "Year style");

        var obj = new ClassWithColumnAttributes(
            id: Guid.NewGuid().ToString(),
            countryOfOrigin: "Germany",
            model: "Golf",
            make: "Volkswagen",
            year: 1992,
            kW: 96,
            length: 428.4m);

        // Act
        await s.AddAsRowAsync(obj, ctx, Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal($"{obj.Year} (leap year)", sheet["A1"].StringValue);
        Assert.Equal(obj.Make[..8], sheet["B1"].StringValue);
        Assert.Equal(obj.CountryOfOrigin, sheet["C1"].StringValue);
        Assert.Equal(obj.Model, sheet["D1"].StringValue);
        Assert.Equal(obj.kW, sheet["E1"].DecimalValue);
        Assert.Equal($"{obj.Length} cm", sheet["F1"].StringValue);
        Assert.Equal(obj.Id, sheet["G1"].StringValue);

        Assert.True(sheet["A1"].Style.Font.Bold);

        Assert.Equal(10, sheet.Column("A").Width, 4);
        Assert.Equal(40, sheet.Column("C").Width, 4);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_StartWorksheet_ColumnWidthFromAttribute(bool createWorksheetOptions)
    {
        // Arrange
        var ctx = ColumnWidthContext.Default.ClassWithColumnWidth;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        if (createWorksheetOptions)
        {
            var worksheetOptions = ctx.CreateWorksheetOptions();
            await s.StartWorksheetAsync("Sheet", worksheetOptions, Token);
        }
        else
        {
            await s.StartWorksheetAsync("Sheet", ctx, Token);
        }

        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(20d, sheet.Column("A").Width, 4);
    }
}
