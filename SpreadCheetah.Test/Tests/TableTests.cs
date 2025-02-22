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
}
