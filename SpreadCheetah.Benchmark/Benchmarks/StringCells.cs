using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System.IO.Packaging;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class StringCells : IDisposable
{
    private const string SheetName = "Book1";

    public Stream Stream { get; set; } = null!;
    public IList<IList<string>> Values { get; private set; } = null!;

    [Params(20000)]
    public int NumberOfRows { get; set; }

    public int NumberOfColumns { get; set; } = 10;

    [GlobalSetup]
    public void GlobalSetup()
    {
        Values = new List<IList<string>>(NumberOfRows);
        for (var row = 1; row <= NumberOfRows; ++row)
        {
            var cells = new List<string>(NumberOfColumns);
            Values.Add(cells);
            for (var col = 1; col <= NumberOfColumns; ++col)
            {
                cells.Add($"{col}-{row}");
            }
        }
    }

    [IterationSetup]
    public void IterationSetup()
    {
        Stream = new MemoryStream(2_000_000);
    }

    [IterationCleanup]
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposing) return;
        Stream?.Dispose();
        Stream = null!;
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream, options);
        await spreadsheet.StartWorksheetAsync(SheetName);
        var cells = new DataCell[NumberOfColumns];

        for (var row = 0; row < Values.Count; ++row)
        {
            var rowValues = Values[row];
            for (var col = 0; col < rowValues.Count; ++col)
            {
                cells[col] = new DataCell(rowValues[col]);
            }

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }

    [Benchmark]
    public void EpPlus4()
    {
        using var package = new ExcelPackage(Stream) { Compression = CompressionLevel.BestSpeed };
        var worksheet = package.Workbook.Worksheets.Add(SheetName);

        for (var row = 0; row < Values.Count; ++row)
        {
            var rowValues = Values[row];
            for (var col = 0; col < rowValues.Count; ++col)
            {
                worksheet.Cells[row + 1, col + 1].Value = rowValues[col];
            }
        }

        package.Save();
    }

    [Benchmark]
    public void OpenXmlSax()
    {
        using var xl = SpreadsheetDocument.Create(Stream, SpreadsheetDocumentType.Workbook);
        xl.CompressionOption = CompressionOption.SuperFast;
        var workbookPart = xl.AddWorkbookPart();
        var wsp = workbookPart.AddNewPart<WorksheetPart>();

        var oxw = OpenXmlWriter.Create(wsp);
        oxw.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
        oxw.WriteStartElement(new SheetData());

        var rowObject = new Row();
        var cellAttributes = new[] { new OpenXmlAttribute("t", "", "inlineStr") };
        var cell = new OpenXmlCell();
        var inlineString = new InlineString();
        var rowAttributes = new OpenXmlAttribute[1];

        for (var row = 0; row < NumberOfRows; ++row)
        {
            rowAttributes[0] = new OpenXmlAttribute("r", "", (row + 1).ToString());
            oxw.WriteStartElement(rowObject, rowAttributes);
            var rowValues = Values[row];

            for (var col = 0; col < rowValues.Count; ++col)
            {
                oxw.WriteStartElement(cell, cellAttributes);
                oxw.WriteStartElement(inlineString);
                oxw.WriteElement(new Text(rowValues[col]));
                oxw.WriteEndElement();
                oxw.WriteEndElement();
            }

            oxw.WriteEndElement();
        }

        oxw.WriteEndElement();
        oxw.WriteEndElement();
        oxw.Close();

        oxw = OpenXmlWriter.Create(workbookPart);
        oxw.WriteStartElement(new Workbook());
        oxw.WriteStartElement(new Sheets());

        oxw.WriteElement(new Sheet()
        {
            Name = "Sheet1",
            SheetId = 1,
            Id = workbookPart.GetIdOfPart(wsp)
        });

        oxw.WriteEndElement();
        oxw.WriteEndElement();
        oxw.Close();
    }

    [Benchmark]
    public void OpenXmlDom()
    {
        using var xl = SpreadsheetDocument.Create(Stream, SpreadsheetDocumentType.Workbook);
        xl.CompressionOption = CompressionOption.SuperFast;
        var workbookpart = xl.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();
        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
        worksheetPart.Worksheet.AppendChild(sheetData);

        var sheets = workbookpart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = workbookpart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = SheetName
        };

        var cells = new OpenXmlCell[NumberOfColumns];

        for (var row = 0; row < Values.Count; ++row)
        {
            var rowValues = Values[row];
            for (var col = 0; col < rowValues.Count; ++col)
            {
                var inlineString = new InlineString();
                inlineString.AppendChild(new Text(rowValues[col]));
                var cell = new OpenXmlCell { DataType = OpenXmlCellValue.InlineString };
                cell.AppendChild(inlineString);
                cells[col] = cell;
            }

            var rowObject = new Row(cells) { RowIndex = (uint)row + 1 };
            sheetData.AppendChild(rowObject);
        }

        sheets.AppendChild(sheet);
        workbookpart.Workbook.Save();
    }

    [Benchmark]
    public void ClosedXml()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(SheetName);

        for (var row = 0; row < Values.Count; ++row)
        {
            var rowValues = Values[row];
            for (var col = 0; col < rowValues.Count; ++col)
            {
                worksheet.Cell(row + 1, col + 1).SetValue(rowValues[col]);
            }
        }

        workbook.SaveAs(Stream);
    }
}
