using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SpreadCheetah.Benchmark.Benchmarks
{
    [SimpleJob(RuntimeMoniker.Net48)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    [SimpleJob(RuntimeMoniker.Net50)]
    [MemoryDiagnoser]
    public class MixedDataTypeCells
    {
        private List<string> _stringValues1 = null!;
        private List<string> _stringValues2 = null!;

        [Params(100000)]
        public int NumberOfRows { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            _stringValues1 = new List<string>(NumberOfRows);
            _stringValues2 = new List<string>(NumberOfRows);

            for (var row = 0; row < NumberOfRows; ++row)
            {
                _stringValues1.Add("Ola-" + row);
                _stringValues2.Add(row + "-Nordmann");
            }
        }

        [Benchmark]
        public async Task SpreadCheetah()
        {
            await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
            await spreadsheet.StartWorksheetAsync("Book1");

            var cells = new DataCell[10];

            for (var row = 0; row < NumberOfRows; ++row)
            {
                var even = row % 2 == 0;
                cells[0] = new DataCell(row);
                cells[1] = new DataCell(_stringValues1[row]);
                cells[2] = new DataCell(_stringValues2[row]);
                cells[3] = new DataCell(1950 + row / 1000);
                cells[4] = new DataCell(5.67 + row % 10);
                cells[5] = new DataCell(even);
                cells[6] = new DataCell(even ? "Norway" : "Sweden");
                cells[7] = new DataCell(!even);
                cells[8] = new DataCell(0.991f + row / 10000.0);
                cells[9] = new DataCell(even ? -23 : 23);

                await spreadsheet.AddRowAsync(cells);
            }

            await spreadsheet.FinishAsync();
        }
    }
}
