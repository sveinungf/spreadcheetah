using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Threading.Tasks;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.Benchmark.Benchmarks
{
    [SimpleJob(RuntimeMoniker.Net48)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    [SimpleJob(RuntimeMoniker.NetCoreApp50)]
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
            await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream);
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
            xl.AddWorkbookPart();
            var wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

            var oxw = OpenXmlWriter.Create(wsp);
            oxw.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
            oxw.WriteStartElement(new SheetData());

            var rowObject = new Row();
            var cellAttributes = new[] { new OpenXmlAttribute("t", null, "inlineStr") };
            var cell = new OpenXmlCell();
            var inlineString = new InlineString();

            for (var row = 0; row < NumberOfRows; ++row)
            {
                var rowAttributes = new[] { new OpenXmlAttribute("r", null, (row + 1).ToString()) };
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
            using var xl = SpreadsheetDocument.Create(Stream, SpreadsheetDocumentType.Workbook);
            xl.CompressionOption = CompressionOption.SuperFast;
            var workbookpart = xl.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
            worksheetPart.Worksheet.AppendChild(sheetData);

            var sheets = xl.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet
            {
                Id = xl.WorkbookPart.GetIdOfPart(worksheetPart),
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
                    var cell = new OpenXmlCell { DataType = CellValues.InlineString };
                    cell.AppendChild(inlineString);
                    cells[col] = cell;
                }

                var rowObject = new Row(cells) { RowIndex = (uint)row + 1 };
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
}
