using SpreadCheetah.Tables;
using SpreadCheetah.TestHelpers.Assertions;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetTableTests
{
    [Fact]
    public async Task Spreadsheet_Table_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];
        DataCell[] dataRow1 = [new("Ford"), new("Mondeo"), new(1993)];
        DataCell[] dataRow2 = [new("Volkswagen"), new("Polo"), new(1975)];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync(dataRow1);
        await spreadsheet.AddRowAsync(dataRow2);
        await spreadsheet.FinishTableAsync();
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(headerNames, actualTable.Columns.Select(x => x.Name));
        Assert.Equal("Table1", actualTable.Name);
        Assert.Equal("TableStyleLight1", actualTable.TableStyle);
        Assert.Equal("A1:C3", actualTable.CellRangeReference);
        Assert.True(actualTable.ShowAutoFilter);
        Assert.True(actualTable.ShowHeaderRow);
        Assert.True(actualTable.BandedRows);
        Assert.False(actualTable.BandedColumns);
        Assert.False(actualTable.ShowTotalRow);
    }

    [Fact]
    public async Task Spreadsheet_Table_TotalRow()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light19);
        string[] headerNames = ["Make", "Model", "Price"];
        const string totalRowLabel = "Average price";
        const TableTotalRowFunction totalRowFunction = TableTotalRowFunction.Average;
        table.Column(1).TotalRowLabel = totalRowLabel;
        table.Column(3).TotalRowFunction = totalRowFunction;
        DataCell[] dataRow1 = [new("Ford"), new("Mondeo"), new(190000)];
        DataCell[] dataRow2 = [new("Volkswagen"), new("Polo"), new(150000)];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync(dataRow1);
        await spreadsheet.AddRowAsync(dataRow2);
        await spreadsheet.FinishTableAsync();
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var actualColumns = actualTable.Columns;
        Assert.Equal(headerNames, actualColumns.Select(x => x.Name));
        Assert.Equal([totalRowLabel, null, null], actualColumns.Select(x => x.TotalRowLabel));
        Assert.Equal([null, null, totalRowFunction], actualColumns.Select(x => x.TotalRowFunction));
        Assert.Equal("Table1", actualTable.Name);
        Assert.Equal("TableStyleLight19", actualTable.TableStyle);
        Assert.Equal("A1:C4", actualTable.CellRangeReference);
        Assert.True(actualTable.ShowAutoFilter);
        Assert.True(actualTable.ShowHeaderRow);
        Assert.True(actualTable.ShowTotalRow);
        Assert.True(actualTable.BandedRows);
        Assert.False(actualTable.BandedColumns);
    }
}
