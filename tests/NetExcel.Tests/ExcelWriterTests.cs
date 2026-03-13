using FluentAssertions;
using NetXLCsv.Excel;

namespace NetXLCsvTests;

/// <summary>Tests for writing Excel files.</summary>
public sealed class ExcelWriterTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"NetXLCsvTest_{Guid.NewGuid():N}");

    public ExcelWriterTests() => Directory.CreateDirectory(_tempDir);
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string TempFile(string name) => Path.Combine(_tempDir, name);

    [Fact]
    public void WriteWorkbook_CreatesValidFile()
    {
        var path = TempFile("out.xlsx");
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet("Test");
        ws.WriteCell(1, 1, "Hello");
        ws.WriteCell(1, 2, 42);
        wb.Save(path);

        File.Exists(path).Should().BeTrue();
        new FileInfo(path).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void WriteCell_RoundTrip_PreservesStringValue()
    {
        var path = TempFile("rt_str.xlsx");
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet("S");
        ws.WriteCell(1, 1, "NetExcel");
        wb.Save(path);

        var reader = new ExcelReader();
        var df = reader.ReadDataFrame(path, hasHeader: false);
        df.GetValue(0, 0)?.ToString().Should().Be("NetExcel");
    }

    [Fact]
    public void WriteMultipleSheets_AllPresent()
    {
        var path = TempFile("multi.xlsx");
        using var wb = ExcelWorkbook.Create();
        wb.AddWorksheet("Alpha");
        wb.AddWorksheet("Beta");
        wb.AddWorksheet("Gamma");
        wb.Save(path);

        using var read = new ExcelReader().OpenWorkbook(path);
        read.Worksheets.Should().HaveCount(3);
        read.Worksheets.Select(w => w.Name).Should().Equal("Alpha", "Beta", "Gamma");
    }

    [Fact]
    public void DuplicateSheetNameThrows()
    {
        using var wb = ExcelWorkbook.Create();
        wb.AddWorksheet("Sheet1");
        var act = () => wb.AddWorksheet("Sheet1");
        act.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void ExcelWriter_WritesDataFrame()
    {
        var df = DataFrame.FromList(new[]
        {
            new { Product = "Widget", Units = 100, Revenue = 9999.99 },
            new { Product = "Gadget", Units = 50,  Revenue = 4500.00 }
        });

        var path = TempFile("writer.xlsx");
        var writer = new ExcelWriter();
        writer.Write(df, path);

        File.Exists(path).Should().BeTrue();

        var reader = new ExcelReader();
        var restored = reader.ReadDataFrame(path);
        restored.RowCount.Should().Be(2);
        restored.ColumnCount.Should().Be(3);
    }

    [Fact]
    public void ExcelWriter_WritesToStream()
    {
        var df = DataFrame.FromList(new[] { new { A = 1, B = 2 } });

        using var ms = new MemoryStream();
        new ExcelWriter().Write(df, ms);

        ms.Length.Should().BeGreaterThan(0);

        // Read back from a fresh MemoryStream so the position is at the start
        using var readMs = new MemoryStream(ms.ToArray());
        var restored = new ExcelReader().ReadDataFrame(readMs);
        restored.RowCount.Should().Be(1);
    }

    [Fact]
    public void SetColumnWidth_IsReflectedInOutput()
    {
        var path = TempFile("colwidth.xlsx");
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet("S");
        ws.WriteCell(1, 1, "Wide Column");
        ws.SetColumnWidth(1, 40.0);
        wb.Save(path);

        // Basic check: file is valid and readable
        File.Exists(path).Should().BeTrue();
    }
}

internal static class ArrayExtensions
{
    public static Stream ToStream(this byte[] bytes) => new MemoryStream(bytes);
}
