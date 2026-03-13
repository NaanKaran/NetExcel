using FluentAssertions;
using NetXLCsv.Csv;
using NetXLCsv.Excel;

namespace NetXLCsvTests;

/// <summary>
/// End-to-end integration tests combining DataFrame → Excel → CSV round-trips,
/// and the typical "sales report" workflow described in the specification.
/// </summary>
public sealed class IntegrationTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"NetXLCsvInteg_{Guid.NewGuid():N}");

    public IntegrationTests() => Directory.CreateDirectory(_tempDir);
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string TempFile(string name) => Path.Combine(_tempDir, name);

    // ── Sales report workflow ─────────────────────────────────────────────────

    [Fact]
    public void SalesReport_FilterHighRevenue_ExportToExcel()
    {
        // Arrange: build a "sales" DataFrame in memory
        var salesData = Enumerable.Range(1, 20).Select(i => new
        {
            Region = i % 2 == 0 ? "North" : "South",
            Product = $"Product{i}",
            Revenue = (decimal)(i * 5000)
        });

        var df = DataFrame.FromList(salesData);

        // Act: filter high-revenue rows, select subset, export
        var resultPath = TempFile("report.xlsx");
        df.Filter(r => (decimal)r["Revenue"]! > 50000)
          .Select("Region", "Revenue")
          .ToExcel(resultPath);

        // Assert: file exists and has correct content
        File.Exists(resultPath).Should().BeTrue();

        var reader = new ExcelReader();
        var result = reader.ReadDataFrame(resultPath);
        result.RowCount.Should().BeGreaterThan(0);
        result.ColumnCount.Should().Be(2);
        result.Schema.Columns.Select(c => c.Name).Should().Equal("Region", "Revenue");
    }

    [Fact]
    public void CsvRoundTrip_PreservesAllRows()
    {
        var original = DataFrame.FromList(Enumerable.Range(1, 50).Select(i =>
            new { Id = i, Name = $"Person{i}", Score = i * 1.5 }));

        var path = TempFile("roundtrip.csv");
        original.ToCsv(path);

        var restored = DataFrame.ReadCsv(path);
        restored.RowCount.Should().Be(50);
        restored.ColumnCount.Should().Be(3);
    }

    [Fact]
    public void ExcelRoundTrip_HeadersAndValuesPreserved()
    {
        var original = DataFrame.FromList(new[]
        {
            new { Name = "Alice",  Department = "Engineering", Salary = 95000m },
            new { Name = "Bob",    Department = "Marketing",   Salary = 72000m },
            new { Name = "Carol",  Department = "Engineering", Salary = 88000m }
        });

        var path = TempFile("employees.xlsx");
        original.ToExcel(path);

        var restored = DataFrame.ReadExcel(path);
        restored.RowCount.Should().Be(3);
        restored.ColumnCount.Should().Be(3);
        restored.Schema.Columns.Select(c => c.Name).Should().Equal("Name", "Department", "Salary");
    }

    [Fact]
    public void AddColumn_ThenFilter_ThenExport_ProducesCorrectOutput()
    {
        var df = DataFrame.FromList(new[]
        {
            new { Region = "APAC", Revenue = 120000m },
            new { Region = "EMEA", Revenue = 45000m  },
            new { Region = "APAC", Revenue = 80000m  }
        });

        var path = TempFile("enriched.xlsx");
        df.AddColumn("Source", "FY2025")
          .Filter(r => (decimal)r["Revenue"]! > 50000)
          .ToExcel(path);

        var result = new ExcelReader().ReadDataFrame(path);
        result.RowCount.Should().Be(2);          // EMEA filtered out
        result.Schema.HasColumn("Source").Should().BeTrue();
    }

    [Fact]
    public void GroupBy_AggregateCounts_CorrectPerGroup()
    {
        var df = DataFrame.FromList(Enumerable.Range(1, 100).Select(i =>
            new { Category = $"Cat{i % 5}", Value = i }));

        var groups = df.GroupBy("Category");

        groups.Should().HaveCount(5);
        foreach (var g in groups.Values)
            g.RowCount.Should().Be(20);
    }

    [Fact]
    public void MultiSheetWorkbook_EachSheetHasCorrectData()
    {
        var path = TempFile("multis.xlsx");
        using var wb = ExcelWorkbook.Create();

        var sheet1 = (ExcelWorksheet)wb.AddWorksheet("Employees");
        sheet1.WriteTable(1, 1, DataFrame.FromList(new[]
        {
            new { Name = "Alice", Dept = "Eng" },
            new { Name = "Bob",   Dept = "HR"  }
        }));

        var sheet2 = (ExcelWorksheet)wb.AddWorksheet("Products");
        sheet2.WriteTable(1, 1, DataFrame.FromList(new[]
        {
            new { Sku = "W001", Price = 9.99 }
        }));

        wb.Save(path);

        // Verify both sheets are readable
        var reader = new ExcelReader();
        var emp = reader.ReadDataFrame(path, sheetName: "Employees");
        var prod = reader.ReadDataFrame(path, sheetName: "Products");

        emp.RowCount.Should().Be(2);
        prod.RowCount.Should().Be(1);
    }
}
