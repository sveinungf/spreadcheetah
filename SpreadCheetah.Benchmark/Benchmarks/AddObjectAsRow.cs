using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Benchmark.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SpreadCheetah.Benchmark.Benchmarks
{
    [SimpleJob(RuntimeMoniker.Net50)]
    [MemoryDiagnoser]
    public class AddObjectAsRow
    {
        public IList<Student> Students { get; private set; } = null!;

        [Params(10, 1000, 100000)]
        public int NumberOfRows { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            Students = StudentGenerator.Generate(NumberOfRows);
        }

        [Benchmark]
        public async Task SpreadCheetah()
        {
            await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
            await spreadsheet.StartWorksheetAsync("Book1");

            for (var i = 0; i < Students.Count; ++i)
            {
                await spreadsheet.AddAsRowAsync(Students[i]);
            }

            await spreadsheet.FinishAsync();
        }
    }
}
