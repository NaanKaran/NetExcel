using FluentAssertions;
using NetXLCsv.Excel;
using NetXLCsv.Streaming;

namespace NetXLCsvTests;

/// <summary>Tests for streaming writers.</summary>
public sealed class StreamingTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"NetXLCsvTest_{Guid.NewGuid():N}");

    public StreamingTests() => Directory.CreateDirectory(_tempDir);
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string TempFile(string name) => Path.Combine(_tempDir, name);

    // ── StreamingExcelWriter ──────────────────────────────────────────────────

    [Fact]
    public void StreamingExcel_WritesHeaderAndDataRows()
    {
        var path = TempFile("stream.xlsx");
        using (var writer = StreamingExcelWriter.Create(path))
        {
            writer.WriteHeader("Name", "Age", "Country");
            writer.WriteRow("Alice", 30, "India");
            writer.WriteRow("Bob",   25, "UK");
        }

        File.Exists(path).Should().BeTrue();
        var df = new ExcelReader().ReadDataFrame(path);
        df.RowCount.Should().Be(2);
        df.ColumnCount.Should().Be(3);
    }

    [Fact]
    public void StreamingExcel_RowsWrittenCountIsCorrect()
    {
        var path = TempFile("count.xlsx");
        using var writer = StreamingExcelWriter.Create(path);
        writer.WriteHeader("A", "B");
        for (int i = 0; i < 100; i++)
            writer.WriteRow(i, i * 2);

        writer.RowsWritten.Should().Be(101); // 1 header + 100 data
    }

    [Fact]
    public void StreamingExcel_1000Rows_CompletesWithinReasonableTime()
    {
        var path = TempFile("perf.xlsx");
        var sw = System.Diagnostics.Stopwatch.StartNew();

        using (var writer = StreamingExcelWriter.Create(path))
        {
            writer.WriteHeader("Id", "Name", "Value");
            for (int i = 0; i < 1_000; i++)
                writer.WriteRow(i, $"Row{i}", i * 3.14);
        }

        sw.Stop();
        sw.ElapsedMilliseconds.Should().BeLessThan(5000, "1k rows should write quickly");
    }

    [Fact]
    public void StreamingExcel_WriteAfterDispose_Throws()
    {
        var path = TempFile("disposed.xlsx");
        var writer = StreamingExcelWriter.Create(path);
        writer.WriteHeader("A");
        writer.Dispose();

        var act = () => writer.WriteRow("x");
        act.Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void StreamingExcel_SupportsNullValues()
    {
        var path = TempFile("nulls.xlsx");
        using (var writer = StreamingExcelWriter.Create(path))
        {
            writer.WriteHeader("A", "B");
            writer.WriteRow(null, "hello");
            writer.WriteRow(42, null);
        }

        File.Exists(path).Should().BeTrue();
    }

    // ── StreamingCsvWriter ────────────────────────────────────────────────────

    [Fact]
    public void StreamingCsv_WritesHeaderAndDataRows()
    {
        var path = TempFile("stream.csv");
        using (var writer = StreamingCsvWriter.Create(path))
        {
            writer.WriteHeader("Name", "Age");
            writer.WriteRow("Alice", 30);
            writer.WriteRow("Bob",   25);
        }

        var lines = File.ReadAllLines(path);
        lines.Should().HaveCount(3);
        lines[0].Should().Be("Name,Age");
        lines[1].Should().Contain("Alice");
    }

    [Fact]
    public void StreamingCsv_RowCountIsTracked()
    {
        var path = TempFile("csvcount.csv");
        using var writer = StreamingCsvWriter.Create(path);
        for (int i = 0; i < 50; i++)
            writer.WriteRow(i, $"Item{i}");

        writer.RowsWritten.Should().Be(50);
    }

    [Fact]
    public void StreamingCsv_QuotesFieldsWithCommasAndNewlines()
    {
        var path = TempFile("quote.csv");
        using (var writer = StreamingCsvWriter.Create(path))
        {
            writer.WriteHeader("Note");
            writer.WriteRow("Hello, World");
        }

        var content = File.ReadAllText(path);
        content.Should().Contain("\"Hello, World\"");
    }

    [Fact]
    public void StreamingCsv_WriteAfterDispose_Throws()
    {
        var path = TempFile("disp.csv");
        var writer = StreamingCsvWriter.Create(path);
        writer.Dispose();
        var act = () => writer.WriteRow("x");
        act.Should().Throw<ObjectDisposedException>();
    }
}
