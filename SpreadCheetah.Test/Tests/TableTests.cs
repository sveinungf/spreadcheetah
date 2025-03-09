using SpreadCheetah.Tables;

namespace SpreadCheetah.Test.Tests;

public static class TableTests
{
    [Fact]
    public static void Table_Create_InvalidStyle()
    {
        // Arrange
        const TableStyle style = TableStyle.Dark11 + 1;

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => new Table(style));
    }

    [Fact]
    public static void Table_Create_InvalidTotalRowFunction()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);
        var column = table.Column(1);
        const TableTotalRowFunction function = default;

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => column.TotalRowFunction = function);
    }

    [Fact]
    public static void Table_Create_InvalidNumberOfColumns()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => table.NumberOfColumns = 0);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("cr")]
    [InlineData("CR")]
    [InlineData("Table")]
    [InlineData("Table1")]
    [InlineData("Table.1")]
    [InlineData("Table\\1")]
    [InlineData("A")]
    [InlineData("b")]
    [InlineData("_")]
    [InlineData("\\")]
    [InlineData("A.1")]
    [InlineData("AAAA1")]
    public static void Table_Create_ValidName(string? name)
    {
        // Act
        var exception = Record.Exception(() => new Table(TableStyle.Light1, name));

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("c")]
    [InlineData("C")]
    [InlineData("r")]
    [InlineData("R")]
    [InlineData("Table 1")]
    [InlineData("1Table")]
    [InlineData(".Table")]
    [InlineData("A1")]
    [InlineData("a1")]
    [InlineData("dog2")]
    [InlineData("R1C1")]
    [InlineData("r1c1")]
    [InlineData("C10")]
    [InlineData("R10")]
    [InlineData("C11\\")]
    [InlineData("R123abc")]
    public static void Table_Create_InvalidName(string name)
    {
        // Act
        var exception = Record.Exception(() => new Table(TableStyle.Light1, name));

        // Assert
        Assert.NotNull(exception);
    }

    [Theory]
    [InlineData(255, true)]
    [InlineData(256, false)]
    public static void Table_Create_LongName(int length, bool valid)
    {
        // Arrange
        var name = new string('a', length);

        // Act
        var exception = Record.Exception(() => new Table(TableStyle.Light1, name));

        // Assert
        Assert.Equal(valid, exception is null);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(16385)]
    public static void Table_Column_InvalidColumnNumber(int columnNumber)
    {
        // Arrange
        var table = new Table(TableStyle.Light1);

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Column(columnNumber));
    }

    [Fact]
    public static void Table_Column_AssignTotalRowLabelWhenHasFunction()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);
        table.Column(1).TotalRowFunction = TableTotalRowFunction.Sum;

        // Act & Assert
        Assert.Throws<SpreadCheetahException>(() => table.Column(1).TotalRowLabel = "Sum");
    }

    [Fact]
    public static void Table_Column_AssignTotalRowFunctionWhenHasLabel()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);
        table.Column(1).TotalRowLabel = "Sum";

        // Act & Assert
        Assert.Throws<SpreadCheetahException>(() => table.Column(1).TotalRowFunction = TableTotalRowFunction.Sum);
    }

    [Fact]
    public static void Table_Column_AssignNullToTotalRowProperties()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);
        table.Column(1).TotalRowLabel = "Sum";
        table.Column(2).TotalRowFunction = TableTotalRowFunction.Count;

        // Act
        var exception1 = Record.Exception(() => table.Column(1).TotalRowFunction = null);
        var exception2 = Record.Exception(() => table.Column(1).TotalRowLabel = null);

        // Assert
        Assert.Null(exception1);
        Assert.Null(exception2);
    }

    [Fact]
    public static void Table_Column_GreaterThanNumberOfColumns()
    {
        // Arrange
        var table = new Table(TableStyle.Light1)
        {
            NumberOfColumns = 3
        };

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Column(4));
    }

    [Fact]
    public static void Table_NumberOfColumns_LessThanGreatestColumnNumber()
    {
        // Arrange
        var table = new Table(TableStyle.Light1);
        table.Column(4).TotalRowFunction = TableTotalRowFunction.Sum;

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => table.NumberOfColumns = 3);
    }
}
