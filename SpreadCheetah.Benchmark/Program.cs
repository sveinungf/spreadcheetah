using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

var baseJob = Job.Default.WithCustomBuildConfiguration("Benchmark");

var config = DefaultConfig.Instance
    .AddJob(baseJob.WithRuntime(CoreRuntime.Core10_0));

BenchmarkSwitcher
    .FromAssembly(typeof(Program).Assembly)
    .Run(args, config);
