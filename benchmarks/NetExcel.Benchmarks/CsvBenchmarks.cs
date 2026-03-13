using BenchmarkDotNet.Attributes;
using NetXLCsv.Csv;
using NetXLCsv.DataFrame;
using NetXLCsv.Streaming;

namespace NetXLCsv.Benchmarks;

/// <summary>
/// Benchmarks for CSV read/write at scale.
/// </summary>
[SimpleJob] // targets the host runtime automatically (net10.0)
[MemoryDiagnoser]
[MinColumn, MaxColumn, MeanColumn, MedianColumn]
public class CsvBenchmarks
{
    [Params(100_000, 1_000_000)]
    public int RowCount { get; set; }

    private string _csvPath = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        _csvPath = Path.Combine(Path.GetTempPath(), $"bench_{Guid.NewGuid():N}.csv");

        // Pre-generate CSV for read benchmarks
        using var writer = StreamingCsvWriter.Create(_csvPath);
        writer.WriteHeader("Id", "Name", "Value", "Category");
        for (int i = 0; i < RowCount; i++)
            writer.WriteRow(i, $"Item{i}", i * 1.23, $"Cat{i % 10}");
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_csvPath)) File.Delete(_csvPath);
    }

    [Benchmark(Baseline = true)]
    public int NetXLCsv_ReadCsv()
    {
        var df = new CsvReader().Read(_csvPath);
        return df.RowCount;
    }

    [Benchmark]
    public int NetXLCsv_StreamRows()
    {
        var reader = new CsvReader();
        int count = 0;
        foreach (var _ in reader.StreamRows(_csvPath, skipHeader: true))
            count++;
        return count;
    }

    [Benchmark]
    public void NetXLCsv_WriteCsv()
    {
        var tempOut = Path.Combine(Path.GetTempPath(), $"out_{Guid.NewGuid():N}.csv");
        try
        {
            using var writer = StreamingCsvWriter.Create(tempOut);
            writer.WriteHeader("Id", "Name", "Value");
            for (int i = 0; i < RowCount; i++)
                writer.WriteRow(i, $"Name{i}", i * 0.5);
        }
        finally
        {
            if (File.Exists(tempOut)) File.Delete(tempOut);
        }
    }
}
