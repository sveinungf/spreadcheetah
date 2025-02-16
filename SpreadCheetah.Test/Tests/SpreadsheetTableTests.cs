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

    [Fact]
    public async Task Spreadsheet_Table_WithoutColumns()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var table = new Table(TableStyle.Light1);

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddRowAsync([new("Hello"), new("World!")]);
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("no columns", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_EmptyHeaderRow(bool numberOfColumnsSet)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1)
        {
            NumberOfColumns = numberOfColumnsSet ? 3 : null
        };

        // Act
        spreadsheet.StartTable(table);
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync([]).AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Equal("Table must have at least one header name.", exception.Message);
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

    [Fact]
    public async Task Spreadsheet_Table_TwoActiveTables()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var table1 = new Table(TableStyle.Light1);
        var table2 = new Table(TableStyle.Light2);

        // Act
        spreadsheet.StartTable(table1);
        var exception = Record.Exception(() => spreadsheet.StartTable(table2));

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("active", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_FinishTableWithoutStartingIt()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        await spreadsheet.AddHeaderRowAsync(["Make", "Model", "Year"]);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("active", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_FinishTableTwice()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(["Make", "Model", "Year"]);
        await spreadsheet.AddRowAsync([new("Ford"), new("Mondeo"), new(1993)]);
        await spreadsheet.FinishTableAsync();
        var exception = await Record.ExceptionAsync(() => spreadsheet.FinishTableAsync().AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("active", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_AddHeaderRowTwice()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames1 = ["Make", "Model", "Year"];
        string[] headerNames2 = ["Width", "Height", "Length"];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames1);
        await spreadsheet.AddHeaderRowAsync(headerNames2);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(headerNames1, actualTable.Columns.Select(x => x.Name));
        Assert.Equal(headerNames1, sheet.Row(1).Select(x => x.StringValue));
        Assert.Equal(headerNames2, sheet.Row(2).Select(x => x.StringValue));
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_MultipleInWorksheet(
        [CombinatorialValues(2, 10, 100)] int count)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        table.Column(2).TotalRowFunction = TableTotalRowFunction.Sum;
        string[] headerNames = ["Text", "Number"];

        // Act
        for (var i = 0; i < count; ++i)
        {
            spreadsheet.StartTable(table);
            await spreadsheet.AddHeaderRowAsync(headerNames);
            await spreadsheet.AddRowAsync([new($"Number {i}"), new(i)]);
            await spreadsheet.FinishTableAsync();
        }

        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(count, sheet.Tables.Count);
        Assert.All(sheet.Tables, x => Assert.Equal(headerNames, x.Columns.Select(c => c.Name)));
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_Table_DuplicateHeaderNames(bool differentCasing)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames =
        [
            "TheMake",
            "Model",
            differentCasing ? "THEMAKE" : "TheMake"
        ];

        // Act
        spreadsheet.StartTable(table);
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync(headerNames).AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("TheMake", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_HeaderRowWithEmptyString()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["Make", "", "Model"];

        // Act
        spreadsheet.StartTable(table);
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync(headerNames).AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("Header for table column 2 was null or empty.", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_Table_HeaderRowLongerThanTableColumns()
    {
        // Arrange
        const int numberOfColumns = 3;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1) { NumberOfColumns = numberOfColumns };
        string[] headerNames = ["A", "B", "C", "D", "E"];

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(headerNames.Take(numberOfColumns), actualTable.Columns.Select(x => x.Name));
        Assert.Equal(headerNames, sheet.Row(1).Select(x => x.StringValue));
    }

    [Fact]
    public async Task Spreadsheet_Table_HeaderRowShorterThanTableColumns()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1) { NumberOfColumns = 5 };
        string[] headerNames = ["A", "B", "C"];

        // Act
        spreadsheet.StartTable(table);
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync(headerNames).AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Equal("Table was expected to have 5 header names, but only 3 were supplied.", exception.Message);
    }

    [Fact]
    public async Task Spreadsheet_Table_HeaderRowEndingBeforeTableStart()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        string[] headerNames = ["A", "B", "C"];

        // Act
        spreadsheet.StartTable(table, "D");
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync(headerNames).AsTask());

        // Assert
        Assert.IsType<SpreadCheetahException>(exception);
        Assert.Equal("Table must have at least one header name.", exception.Message);
    }

    // TODO: Test for special characters in header names
    // TODO: Test for formula characters in header names
    // TODO: Test for multiple worksheets with tables.
    // TODO: Test for styling on top of table (differential/dxf?)
    // TODO: Test for valid table names
    // TODO: Test for invalid table names
}
