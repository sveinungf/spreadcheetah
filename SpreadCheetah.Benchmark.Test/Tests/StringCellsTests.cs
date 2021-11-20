using ClosedXML.Excel;
using SpreadCheetah.Benchmark.Benchmarks;
using SpreadCheetah.Benchmark.Test.Helpers;
using Xunit;

namespace SpreadCheetah.Benchmark.Test.Tests
{
    public sealed class StringCellsTests : IDisposable
    {
        private const int NumberOfColumns = 10;
        private const int NumberOfRows = 20000;
        private readonly StringCells _stringCells;

        public StringCellsTests()
        {
            _stringCells = new StringCells
            {
                NumberOfColumns = NumberOfColumns,
                NumberOfRows = NumberOfRows,
                Stream = new MemoryStream()
            };

            _stringCells.GlobalSetup();
        }

        public void Dispose()
        {
            _stringCells.Dispose();
        }

        [Fact]
        public async Task StringCells_SpreadCheetah_CorrectCellValues()
        {
            // Act
            await _stringCells.SpreadCheetah();

            // Assert
            AssertCellValuesEqual();
        }

        [Fact]
        public void StringCells_EpPlus4_CorrectCellValues()
        {
            // Act
            _stringCells.EpPlus4();

            // Assert
            AssertCellValuesEqual();
        }

        [Fact]
        public void StringCells_OpenXmlSax_CorrectCellValues()
        {
            // Act
            _stringCells.OpenXmlSax();

            // Assert
            AssertCellValuesEqual();
        }

        [Fact]
        public void StringCells_OpenXmlDom_CorrectCellValues()
        {
            // Act
            _stringCells.OpenXmlDom();

            // Assert
            AssertCellValuesEqual();
        }

        [Fact]
        public void StringCells_ClosedXml_CorrectCellValues()
        {
            // Act
            _stringCells.ClosedXml();

            // Assert
            AssertCellValuesEqual();
        }

        private void AssertCellValuesEqual()
        {
            SpreadsheetAssert.Valid(_stringCells.Stream);

            _stringCells.Stream.Position = 0;
            using var workbook = new XLWorkbook(_stringCells.Stream);
            var worksheet = workbook.Worksheets.Single();
            var values = _stringCells.Values;

            for (var r = 0; r < values.Count; ++r)
            {
                var row = worksheet.Row(r + 1);
                var rowValues = values[r];

                for (var c = 0; c < rowValues.Count; ++c)
                {
                    var cell = row.Cell(c + 1);
                    Assert.Equal(rowValues[c], cell.Value);
                }
            }
        }
    }
}
