using ClosedXML.Excel;
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
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualTable = Assert.Single(worksheet.Tables);
        Assert.Equal(headerNames, actualTable.Fields.Select(x => x.Name));
        Assert.Equal("Table1", actualTable.Name);
        Assert.Equal("TableStyleLight1", actualTable.Theme.Name);
        Assert.Equal("A1:C3", actualTable.RangeAddress.ToString());
        Assert.True(actualTable.ShowAutoFilter);
        Assert.True(actualTable.ShowHeaderRow);
        Assert.True(actualTable.ShowRowStripes);
        Assert.False(actualTable.ShowColumnStripes);
        Assert.False(actualTable.ShowTotalsRow);
    }
}
