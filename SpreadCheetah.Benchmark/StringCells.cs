using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using SpreadCheetah.Benchmark.Helpers;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

#pragma warning disable CA1001 // Using Cleanup instead of Dispose

namespace SpreadCheetah.Benchmark
{
    [SimpleJob(RuntimeMoniker.Net48)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    [MemoryDiagnoser]
    public class StringCells
    {
        private const int NumberOfColumns = 10;
        private const string SheetName = "Book1";

        private List<List<Cell>> _cells = null!;
        private List<string> _epplusRows = null!;
        private List<List<OpenXmlCell>> _openXmlDomCells = null!;
        private List<List<Text>> _openXmlSaxCells = null!;
        private List<List<string>> _stringValues = null!;
        private Stream _stream = null!;

        [Params(20000)]
        public int NumberOfRows { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            _cells = new List<List<Cell>>(NumberOfRows);
            _epplusRows = new List<string>(NumberOfRows);
            _stringValues = new List<List<string>>(NumberOfRows);
            _openXmlSaxCells = new List<List<Text>>(NumberOfRows);

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var columns = Enumerable.Range(1, NumberOfColumns).ToList();
                _cells.Add(columns.Select(x => new Cell($"{x}-{row}")).ToList());
                _epplusRows.Add(string.Join(",", columns.Select(x => $"{x}-{row}")));
                _openXmlSaxCells.Add(columns.Select(x => new Text($"{x}-{row}")).ToList());
                _stringValues.Add(columns.Select(x => $"{x}-{row}").ToList());
            }
        }

        [IterationSetup]
        public void IterationSetup()
        {
            _openXmlDomCells = new List<List<OpenXmlCell>>(NumberOfRows);

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var openXmlCells = Enumerable.Range(1, NumberOfColumns).Select(x => OpenXmlHelper.CreateCell($"{x}-{row}")).ToList();
                _openXmlDomCells.Add(openXmlCells);
            }

            _stream = new MemoryStream(2_000_000);
        }

        [IterationCleanup]
        public void IterationCleanup() => _stream.Dispose();

        [Benchmark]
        public async Task SpreadCheetah()
        {
            using var spreadsheet = await Spreadsheet.CreateNewAsync(_stream);
            await spreadsheet.StartWorksheetAsync(SheetName);
            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var cells = _cells[row - 1];
                await spreadsheet.AddRowAsync(cells);
            }

            await spreadsheet.FinishAsync();
        }

        [Benchmark]
        public void EpPlus4()
        {
            using var package = new ExcelPackage(_stream) { Compression = CompressionLevel.BestSpeed };
            var worksheet = package.Workbook.Worksheets.Add(SheetName);
            var dataTypes = Enumerable.Repeat(eDataTypes.String, NumberOfColumns).ToArray();
            var excelTextFormat = new ExcelTextFormat { DataTypes = dataTypes };

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var cells = worksheet.Cells[row, 1];
                var rowText = _epplusRows[row - 1];
                cells.LoadFromText(rowText, excelTextFormat);
            }

            package.Save();
        }

        [Benchmark]
        public void OpenXmlSax()
        {
            using var xl = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);
            xl.CompressionOption = CompressionOption.SuperFast;
            xl.AddWorkbookPart();
            var wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

            var oxw = OpenXmlWriter.Create(wsp);
            oxw.WriteStartElement(new Worksheet());
            oxw.WriteStartElement(new SheetData());

            var rowObject = new Row();
            var cellAttributes = new[] { new OpenXmlAttribute("t", null, "inlineStr") };
            var cell = new OpenXmlCell();
            var inlineString = new InlineString();

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var rowAttributes = new[] { new OpenXmlAttribute("r", null, row.ToString()) };
                oxw.WriteStartElement(rowObject, rowAttributes);
                var cells = _openXmlSaxCells[row - 1];

                for (var col = 1; col <= NumberOfColumns; ++col)
                {
                    oxw.WriteStartElement(cell, cellAttributes);
                    oxw.WriteStartElement(inlineString);
                    oxw.WriteElement(cells[col - 1]);
                    oxw.WriteEndElement();
                    oxw.WriteEndElement();
                }

                oxw.WriteEndElement();
            }

            oxw.WriteEndElement();
            oxw.WriteEndElement();
            oxw.Close();

            oxw = OpenXmlWriter.Create(xl.WorkbookPart);
            oxw.WriteStartElement(new Workbook());
            oxw.WriteStartElement(new Sheets());

            oxw.WriteElement(new Sheet()
            {
                Name = "Sheet1",
                SheetId = 1,
                Id = xl.WorkbookPart.GetIdOfPart(wsp)
            });

            oxw.WriteEndElement();
            oxw.WriteEndElement();
            oxw.Close();
            xl.Close();
        }

        [Benchmark]
        public void OpenXmlDom()
        {
            using var xl = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);
            xl.CompressionOption = CompressionOption.SuperFast;
            var workbookpart = xl.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.AppendChild(sheetData);

            var sheets = xl.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet
            {
                Id = xl.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = SheetName
            };

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var cells = _openXmlDomCells[row - 1];
                var rowObject = new Row(cells) { RowIndex = (uint)row };
                sheetData.AppendChild(rowObject);
            }

            sheets.AppendChild(sheet);
            workbookpart.Workbook.Save();
            xl.Close();
        }

        [Benchmark]
        public void ClosedXml()
        {
            using var workbook = new XLWorkbook(XLEventTracking.Disabled);
            var worksheet = workbook.Worksheets.Add(SheetName);

            for (var row = 1; row <= NumberOfRows; ++row)
            {
                var columns = _stringValues[row - 1];
                for (var col = 1; col <= NumberOfColumns; ++col)
                {
                    worksheet.Cell(row, col).SetValue(columns[col - 1]);
                }
            }

            workbook.SaveAs(_stream);
        }
    }
}
