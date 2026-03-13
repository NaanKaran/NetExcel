using BenchmarkDotNet.Running;
using NetXLCsv.Benchmarks;

Console.WriteLine("NetXLCsv Benchmarks");
Console.WriteLine("===================");
Console.WriteLine("Run with: dotnet run -c Release");
Console.WriteLine();

// Uncomment the benchmark(s) you want to run:
BenchmarkRunner.Run<ExcelWriteBenchmarks>();
// BenchmarkRunner.Run<CsvBenchmarks>();

// Or run all:
// BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
