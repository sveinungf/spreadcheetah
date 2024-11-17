using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using System.Globalization;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellValueConverterTests
{
    [Fact]
    public async Task CellValueConverter_ClassWithReusedConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithReusedConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithReusedConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("OLA", sheet["A1"].StringValue);
        Assert.Equal("NORDMANN", sheet["C1"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_ClassWithGenericConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithGenericConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithGenericConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("-", sheet["B1"].StringValue);
        Assert.Equal("3.1", sheet["D1"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_TwoClassesUsingTheSameConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj1 = new ClassWithCellValueConverter { Name = "John" };
        var obj2 = new ClassWithReusedConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj1, CellValueConverterContext.Default.ClassWithCellValueConverter);
        await spreadsheet.AddAsRowAsync(obj2, CellValueConverterContext.Default.ClassWithReusedConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("JOHN", sheet["A1"].StringValue);
        Assert.Equal("OLA", sheet["A2"].StringValue);
        Assert.Equal("NORDMANN", sheet["C2"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_ClassWithConverterAndCellStyle()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "ID style");
        var obj = new ClassWithCellValueConverterAndCellStyle { Id = "Abc123" };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithCellValueConverterAndCellStyle);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("ABC123", sheet["A1"].StringValue);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }
    
    [Fact]
    public async Task CellValueConverter_ClassWithConverterOnCustomType()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "PercentType");
        var obj = new ClassWithCellValueConverterOnCustomType
        {
            Property = "Abc123", ComplexProperty = null, PercentType = new Percent(123)
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithCellValueConverterOnCustomType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("Abc123", sheet["A1"].StringValue);
        
        Assert.Equal("-", sheet["B1"].StringValue);
        
        Assert.Equal(123, sheet["C1"].DecimalValue);
        Assert.True(sheet["C1"].Style.Font.Bold);
    }

    [Fact]
    public void CellFormat_ClassWithCellValueConverter_CanReadConverterType()
    {
        // Arrange
        var property = typeof(ClassWithCellValueConverter).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithCellValueConverter.Name), StringComparison.Ordinal));

        // Act
        var cellValueConverterAttr = property?.GetCustomAttribute<CellValueConverterAttribute>();

        // Assert
        Assert.NotNull(property);
        Assert.NotNull(cellValueConverterAttr);
        Assert.Equal(typeof(UpperCaseValueConverter), cellValueConverterAttr.ConverterType);
    }

    [Fact]
    public void CellFormat_ClassWithCellValueConverterAndCellStyle_CanReadConverterType()
    {
        // Arrange
        var property = typeof(ClassWithCellValueConverterAndCellStyle).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithCellValueConverterAndCellStyle.Id), StringComparison.Ordinal));

        // Act
        var cellValueConverterAttr = property?.GetCustomAttribute<CellValueConverterAttribute>();

        // Assert
        Assert.NotNull(property);
        Assert.NotNull(cellValueConverterAttr);
        Assert.Equal(typeof(UpperCaseValueConverter), cellValueConverterAttr.ConverterType);
    }

    [Fact]
    public void CellFormat_ClassWithCellValueConverterOnCustomType_CanReadConverterType()
    {
        // Arrange
        var publicProperties = typeof(ClassWithCellValueConverterOnCustomType).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithCellValueConverterOnCustomType.Property), StringComparison.Ordinal));
        var complexProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithCellValueConverterOnCustomType.ComplexProperty), StringComparison.Ordinal));
        var percentTypeProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithCellValueConverterOnCustomType.PercentType), StringComparison.Ordinal));

        // Act
        var propertyCellValueConverterAttr = property?.GetCustomAttribute<CellValueConverterAttribute>();
        var complexPropertyCellValueConverterAttr = complexProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var percentTypePropertyCellValueConverterAttr = percentTypeProperty?.GetCustomAttribute<CellValueConverterAttribute>();

        // Assert
        Assert.NotNull(property);
        Assert.Null(propertyCellValueConverterAttr);

        Assert.NotNull(complexProperty);
        Assert.NotNull(complexPropertyCellValueConverterAttr);
        Assert.Equal(typeof(NullToDashValueConverter<object?>), complexPropertyCellValueConverterAttr.ConverterType);

        Assert.NotNull(property);
        Assert.NotNull(percentTypePropertyCellValueConverterAttr);
        Assert.Equal(typeof(PercentToNumberConverter), percentTypePropertyCellValueConverterAttr.ConverterType);
    }

    [Fact]
    public void CellFormat_ClassWithGenericConverter_CanReadConverterType()
    {
        // Arrange
        var publicProperties = typeof(ClassWithGenericConverter).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithGenericConverter.FirstName), StringComparison.Ordinal));
        var middleNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithGenericConverter.MiddleName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithGenericConverter.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithGenericConverter.Gpa), StringComparison.Ordinal));

        // Act
        var firstNamePropertyCellValueConverterAttr = firstNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var middleNamePropertyCellValueConverterAttr = middleNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var lastNamePropertyCellValueConverterAttr = lastNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var gpaCellValueConverterAttr = gpaProperty?.GetCustomAttribute<CellValueConverterAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.Null(firstNamePropertyCellValueConverterAttr);

        Assert.NotNull(middleNameProperty);
        Assert.NotNull(middleNamePropertyCellValueConverterAttr);
        Assert.Equal(typeof(NullToDashValueConverter<string>), middleNamePropertyCellValueConverterAttr.ConverterType);

        Assert.NotNull(lastNameProperty);
        Assert.Null(lastNamePropertyCellValueConverterAttr);

        Assert.NotNull(gpaProperty);
        Assert.NotNull(gpaCellValueConverterAttr);
        Assert.Equal(typeof(NullToDashValueConverter<decimal?>), gpaCellValueConverterAttr.ConverterType);
    }

    [Fact]
    public void CellFormat_ClassWithReusedConverter_CanReadConverterType()
    {
        // Arrange
        var publicProperties = typeof(ClassWithReusedConverter).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithReusedConverter.FirstName), StringComparison.Ordinal));
        var middleNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithReusedConverter.MiddleName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithReusedConverter.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithReusedConverter.Gpa), StringComparison.Ordinal));

        // Act
        var firstNamePropertyCellValueConverterAttr = firstNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var middleNamePropertyCellValueConverterAttr = middleNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var lastNamePropertyCellValueConverterAttr = lastNameProperty?.GetCustomAttribute<CellValueConverterAttribute>();
        var gpaCellValueConverterAttr = gpaProperty?.GetCustomAttribute<CellValueConverterAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNamePropertyCellValueConverterAttr);
        Assert.Equal(typeof(UpperCaseValueConverter), firstNamePropertyCellValueConverterAttr.ConverterType);

        Assert.NotNull(middleNameProperty);
        Assert.Null(middleNamePropertyCellValueConverterAttr);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNamePropertyCellValueConverterAttr);
        Assert.Equal(typeof(UpperCaseValueConverter), lastNamePropertyCellValueConverterAttr.ConverterType);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaCellValueConverterAttr);
    }
}
