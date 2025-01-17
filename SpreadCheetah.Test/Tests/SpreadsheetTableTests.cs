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

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_TableStyle(TableStyle tableStyle)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(tableStyle);
        DataCell[] dataRow = [new("Ford"), new("Mondeo"), new(1993)];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(["Make", "Model", "Year"]);
        await spreadsheet.AddRowAsync(dataRow);
        await spreadsheet.FinishTableAsync();
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var expectedStyle = tableStyle is TableStyle.None ? "None" : $"TableStyle{tableStyle}";
        Assert.Equal(expectedStyle, actualTable.TableStyle);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_TotalRow(TableTotalRowFunction function)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light19);
        string[] headerNames = ["Make", "Model", "Price"];
        const string totalRowLabel = "Average price";
        table.Column(1).TotalRowLabel = totalRowLabel;
        table.Column(3).TotalRowFunction = function;
        TableTotalRowFunction? expectedFunction = function is TableTotalRowFunction.None ? null : function;
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
        Assert.Equal([null, null, expectedFunction], actualColumns.Select(x => x.TotalRowFunction));
        Assert.Equal("Table1", actualTable.Name);
        Assert.Equal("TableStyleLight19", actualTable.TableStyle);
        Assert.Equal("A1:C4", actualTable.CellRangeReference);
        Assert.True(actualTable.ShowAutoFilter);
        Assert.True(actualTable.ShowHeaderRow);
        Assert.True(actualTable.ShowTotalRow);
        Assert.True(actualTable.BandedRows);
        Assert.False(actualTable.BandedColumns);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_WithoutRows(bool rowBefore)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        if (rowBefore)
        {
            DataCell[] dataRow = [new("Ford"), new("Mondeo"), new(190000)];
            await spreadsheet.AddRowAsync(dataRow);
        }

        var table = new Table(TableStyle.Light1) { NumberOfColumns = 1 };

        // Act
        spreadsheet.StartTable(table);
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("no rows", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_WithoutColumns(bool hasEmptyHeaderRow)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var table = new Table(TableStyle.Light1);

        // Act
        spreadsheet.StartTable(table);

        if (hasEmptyHeaderRow)
            await spreadsheet.AddHeaderRowAsync([]);

        await spreadsheet.AddRowAsync([new DataCell("Hello"), new DataCell("World!")]);
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("no columns", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    // TODO: Test for table with only header row. Should be allowed.
    // TODO: Test for table with only total row. Should not be allowed.
    // TODO: Test for table with only a single data row. Should be allowed?
    // TODO: Test for table that doesn't start at row 1
    // TODO: Test for table that doesn't start at column A
    // TODO: Test for having two active tables
    // TODO: Test for finishing table twice
    // TODO: Test for implicit finishing a table with FinishAsync
    // TODO: Test for calling AddHeaderRow twice. Only the first call should set the headers.
    // TODO: Test for using AddHeaderRow from source generated code
    // TODO: Test for multiple tables in the same worksheet.
    // TODO: Test for multiple worksheets with tables.
    // TODO: Test for styling on top of table (differential/dxf?)
    // TODO: Test for valid table names
    // TODO: Test for invalid table names
    // TODO: Test for column name changes?
    // TODO: Test for generated column names
}
