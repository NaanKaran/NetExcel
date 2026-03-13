using FluentAssertions;
using NetXLCsv.Excel;

namespace NetXLCsvTests;

/// <summary>Tests for reading .xlsx files.</summary>
public sealed class ExcelReaderTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"NetXLCsvTest_{Guid.NewGuid():N}");

    public ExcelReaderTests() => Directory.CreateDirectory(_tempDir);
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string TempFile(string name) => Path.Combine(_tempDir, name);

    private static string CreateSampleExcel(string path)
    {
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet("Data");
        ws.WriteCell(1, 1, "Name");
        ws.WriteCell(1, 2, "Revenue");
        ws.WriteCell(2, 1, "Alice");
        ws.WriteCell(2, 2, 50000.0);
        ws.WriteCell(3, 1, "Bob");
        ws.WriteCell(3, 2, 120000.0);
        wb.Save(path);
        return path;
    }

    [Fact]
    public void ReadDataFrame_ReturnsCorrectRowAndColumnCount()
    {
        var path = TempFile("sample.xlsx");
        CreateSampleExcel(path);

        var reader = new ExcelReader();
        var df = reader.ReadDataFrame(path);

        df.RowCount.Should().Be(2);
        df.ColumnCount.Should().Be(2);
    }

    [Fact]
    public void ReadDataFrame_HeadersMatchFirstRow()
    {
        var path = TempFile("headers.xlsx");
        CreateSampleExcel(path);

        var reader = new ExcelReader();
        var df = reader.ReadDataFrame(path);

        df.Schema.Columns[0].Name.Should().Be("Name");
        df.Schema.Columns[1].Name.Should().Be("Revenue");
    }

    [Fact]
    public void ReadDataFrame_NamedSheet_Works()
    {
        var path = TempFile("multisheet.xlsx");
        using var wb = ExcelWorkbook.Create();
        var ws1 = (ExcelWorksheet)wb.AddWorksheet("Summary");
        ws1.WriteCell(1, 1, "Total");
        ws1.WriteCell(2, 1, "999");

        var ws2 = (ExcelWorksheet)wb.AddWorksheet("Details");
        ws2.WriteCell(1, 1, "Item");
        ws2.WriteCell(2, 1, "Widget");
        wb.Save(path);

        var reader = new ExcelReader();
        var df = reader.ReadDataFrame(path, sheetName: "Details");

        df.GetValue(0, 0)?.ToString().Should().Be("Widget");
    }

    [Fact]
    public void ReadDataFrame_MissingSheetThrows()
    {
        var path = TempFile("miss.xlsx");
        CreateSampleExcel(path);

        var reader = new ExcelReader();
        var act = () => reader.ReadDataFrame(path, sheetName: "NoSuchSheet");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void OpenWorkbook_ReturnsWorkbookWithWorksheets()
    {
        var path = TempFile("wb.xlsx");
        CreateSampleExcel(path);

        var reader = new ExcelReader();
        using var wb = reader.OpenWorkbook(path);

        wb.Worksheets.Should().HaveCount(1);
        wb.Worksheets[0].Name.Should().Be("Data");
    }

    [Fact]
    public void ReadDataFrame_FileNotFoundThrows()
    {
        var reader = new ExcelReader();
        var act = () => reader.ReadDataFrame("/no/such/file.xlsx");
        act.Should().Throw<FileNotFoundException>();
    }
}
