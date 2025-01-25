using SpreadCheetah.Tables;
using SpreadCheetah.Test.Helpers;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetTableTests
{
    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_Success(bool explicitlyFinishTable)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.AddRowAsync([new("Volkswagen"), new("Polo"), new(1975)]);

        if (explicitlyFinishTable)
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

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(["Make", "Model", "Year"]);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.FinishTableAsync();
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var expectedStyle = tableStyle is TableStyle.None ? "None" : $"TableStyle{tableStyle}";
        Assert.Equal(expectedStyle, actualTable.TableStyle);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_TotalRow(TableTotalRowFunction function, bool explicitlyFinishTable)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light19);
        string[] headerNames = ["Make", "Model", "Price"];
        table.Column(1).TotalRowLabel = "Average price";
        table.Column(3).TotalRowFunction = function;
        TableTotalRowFunction? expectedFunction = function is TableTotalRowFunction.None ? null : function;

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(190000)]);
        await spreadsheet.AddRowAsync([new("Volkswagen"), new("Polo"), new(150000)]);

        if (explicitlyFinishTable)
            await spreadsheet.FinishTableAsync();

        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var actualColumns = actualTable.Columns;
        var columns = Enumerable.Range(1, 3).Select(table.Column);
        Assert.Equal(headerNames, actualColumns.Select(x => x.Name));
        Assert.Equal(columns.Select(x => x.TotalRowLabel), actualColumns.Select(x => x.TotalRowLabel));
        Assert.Equal([null, null, expectedFunction], actualColumns.Select(x => x.TotalRowFunction));
        Assert.Equal("Table1", actualTable.Name);
        Assert.Equal("TableStyleLight19", actualTable.TableStyle);
        Assert.Equal("A1:C4", actualTable.CellRangeReference);
        Assert.Equal("Average price", sheet["A4"].StringValue);
        Assert.True(actualTable.ShowAutoFilter);
        Assert.True(actualTable.ShowHeaderRow);
        Assert.True(actualTable.ShowTotalRow);
        Assert.True(actualTable.BandedRows);
        Assert.False(actualTable.BandedColumns);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_WithoutRows(bool rowBefore, bool totalRow)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        if (rowBefore)
            await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(190000)]);

        var table = new Table(TableStyle.Light1) { NumberOfColumns = 1 };

        if (totalRow)
            table.Column(1).TotalRowLabel = "Result";

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

        await spreadsheet.AddRowAsync([new("Hello"), new("World!")]);
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("no columns", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_WithOnlyHeaderRow()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(headerNames, actualTable.Columns.Select(x => x.Name));
        Assert.Equal("A1:C2", actualTable.CellRangeReference);
    }

    [Fact]
    public async Task Spreadsheet_Table_WithOnlyHeaderAndTotalRows()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];
        table.Column(1).TotalRowLabel = "Result";
        table.Column(3).TotalRowFunction = TableTotalRowFunction.Count;

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var actualColumns = actualTable.Columns;
        var columns = Enumerable.Range(1, 3).Select(table.Column);
        Assert.Equal(headerNames, actualTable.Columns.Select(x => x.Name));
        Assert.Equal(columns.Select(x => x.TotalRowLabel), actualColumns.Select(x => x.TotalRowLabel));
        Assert.Equal(columns.Select(x => x.TotalRowFunction), actualColumns.Select(x => x.TotalRowFunction));
        Assert.Equal("A1:C3", actualTable.CellRangeReference);
        Assert.Equal("Result", sheet["A3"].StringValue);
    }

    [Fact]
    public async Task Spreadsheet_Table_WithOnlyDataRow()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1) { NumberOfColumns = 3 };

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(190000)]);
        await spreadsheet.AddRowAsync([new("Ford"), new("Sierra"), new(140000)]);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(["Column1", "Column2", "Column3"], actualTable.Columns.Select(x => x.Name));
        Assert.Equal("A1:C2", actualTable.CellRangeReference);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_StartingAtRow(
        [CombinatorialValues(2, 10, 999)] int startRow)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];

        for (var i = 1; i < startRow; i++)
        {
            await spreadsheet.AddRowAsync([new(i)]);
        }

        var expectedTableReference = $"A{startRow}:C{startRow + 2}";

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.AddRowAsync([new("Volkswagen"), new("Polo"), new(1975)]);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(expectedTableReference, actualTable.CellRangeReference);
        Assert.Equal(headerNames, actualTable.Columns.Select(x => x.Name));

        var actualHeaderNames = sheet.Row(startRow).Select(x => x.StringValue);
        Assert.Equal(headerNames, actualHeaderNames);
        var firstColumnValues = sheet.Column("A").Cells.Skip(startRow).Take(2).Select(x => x.StringValue);
        Assert.Equal(["Ford", "Volkswagen"], firstColumnValues);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_WithTotalRowAndStartingAtRow(
        [CombinatorialValues(2, 10, 999)] int startRow)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "Model", "Year"];
        table.Column(1).TotalRowLabel = "Oldest";
        table.Column(3).TotalRowFunction = TableTotalRowFunction.Minimum;

        for (var i = 1; i < startRow; i++)
        {
            await spreadsheet.AddRowAsync([new(i)]);
        }

        var expectedTableReference = $"A{startRow}:C{startRow + 3}";

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.AddRowAsync([new("Volkswagen"), new("Polo"), new(1975)]);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        var actualColumns = actualTable.Columns;
        var columns = Enumerable.Range(1, 3).Select(table.Column);
        Assert.Equal(expectedTableReference, actualTable.CellRangeReference);
        Assert.Equal(headerNames, actualColumns.Select(x => x.Name));
        Assert.Equal(columns.Select(x => x.TotalRowLabel), actualColumns.Select(x => x.TotalRowLabel));
        Assert.Equal(columns.Select(x => x.TotalRowFunction), actualColumns.Select(x => x.TotalRowFunction));

        var actualHeaderNames = sheet.Row(startRow).Select(x => x.StringValue);
        Assert.Equal(headerNames, actualHeaderNames);
        var firstColumnValues = sheet.Column("A").Cells.Skip(startRow).Take(3).Select(x => x.StringValue);
        Assert.Equal(["Ford", "Volkswagen", "Oldest"], firstColumnValues);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_StartingAtColumn(
        [CombinatorialValues("B", "K", "CB")] string startColumn)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        var columnNumber = ColumnName.Parse(startColumn);
        var firstHeaderCells = Enumerable.Repeat("", columnNumber - 1);
        var firstDataCells = Enumerable.Repeat(new DataCell(), columnNumber - 1);
        string[] headerNames = ["Make", "Model", "Year"];

        var endColumn = SpreadsheetUtility.GetColumnName(columnNumber + 2);
        var expectedTableReference = $"{startColumn}1:{endColumn}3";

        // Act
        spreadsheet.StartTable(table, startColumn);
        await spreadsheet.AddHeaderRowAsync([.. firstHeaderCells, .. headerNames]);
        await spreadsheet.AddRowAsync([.. firstDataCells, new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.AddRowAsync([.. firstDataCells, new("Volkswagen"), new("Polo"), new(1975)]);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(expectedTableReference, actualTable.CellRangeReference);
        Assert.Equal(headerNames, actualTable.Columns.Select(x => x.Name));

        var actualHeaderNames = sheet.Row(1).Select(x => x.StringValue);
        Assert.Equal(headerNames, actualHeaderNames);
        var firstColumnValues = sheet.Column(startColumn).Cells.Skip(1).Take(2).Select(x => x.StringValue);
        Assert.Equal(["Ford", "Volkswagen"], firstColumnValues);
    }

    // TODO: Test for having two active tables
    // TODO: Test for finishing table twice
    // TODO: Test for calling AddHeaderRow twice. Only the first call should set the headers.
    // TODO: Test for using AddHeaderRow from source generated code
    // TODO: Test for multiple tables in the same worksheet.
    // TODO: Test for multiple worksheets with tables.
    // TODO: Test for styling on top of table (differential/dxf?)
    // TODO: Test for valid table names
    // TODO: Test for invalid table names
    // TODO: Test for column name changes?
    // TODO: Test for generated column names
    // TODO: Test for header row where the header name is null/empty for one of the columns
    // TODO: Test for header row with duplicate header names
}
