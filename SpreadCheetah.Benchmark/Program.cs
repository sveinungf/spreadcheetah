using BenchmarkDotNet.Running;

namespace SpreadCheetah.Benchmark
{
    public static partial class Program
    {
        public static void Main()
        {
            _ = BenchmarkRunner.Run<StringCells>();
        }
    }
}
