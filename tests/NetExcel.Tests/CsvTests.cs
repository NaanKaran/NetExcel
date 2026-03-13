using System.Text;
using FluentAssertions;
using NetXLCsv.Csv;

namespace NetXLCsvTests;

/// <summary>Tests for CSV reading and writing.</summary>
public sealed class CsvTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"NetXLCsvTest_{Guid.NewGuid():N}");

    public CsvTests() => Directory.CreateDirectory(_tempDir);

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private string TempFile(string name) => Path.Combine(_tempDir, name);

    // ── Read ───────────────────────────────────────────────────────────────────

    [Fact]
    public void Read_SimpleCSV_ReturnsCorrectShape()
    {
        var csv = "Name,Age,Country\nAlice,30,India\nBob,25,UK\n";
        var path = TempFile("simple.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var df = reader.Read(path);

        df.RowCount.Should().Be(2);
        df.ColumnCount.Should().Be(3);
    }

    [Fact]
    public void Read_Headerless_UsesGeneratedHeaders()
    {
        var csv = "Alice,30\nBob,25\n";
        var path = TempFile("noheader.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var df = reader.Read(path, hasHeader: false);

        df.ColumnCount.Should().Be(2);
        df.Schema.Columns[0].Name.Should().Be("Col1");
        df.Schema.Columns[1].Name.Should().Be("Col2");
    }

    [Fact]
    public void Read_QuotedFields_HandledCorrectly()
    {
        var csv = "Name,Note\n\"Smith, John\",\"He said \"\"hello\"\"\"\n";
        var path = TempFile("quoted.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var df = reader.Read(path);

        df.RowCount.Should().Be(1);
        df.GetValue(0, 0).Should().Be("Smith, John");
        df.GetValue(0, 1)!.ToString().Should().Contain("hello");
    }

    [Fact]
    public void Read_TypeInference_ParsesIntegers()
    {
        var csv = "Id,Value\n1,100\n2,200\n";
        var path = TempFile("types.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var df = reader.Read(path);

        df.GetValue(0, 0).Should().BeOfType<long>();
        df.GetValue(0, 1).Should().BeOfType<long>();
    }

    [Fact]
    public void Read_CustomDelimiter_SemiColon()
    {
        var csv = "Name;Age\nAlice;30\n";
        var path = TempFile("semi.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var df = reader.Read(path, delimiter: ';');

        df.ColumnCount.Should().Be(2);
        df.RowCount.Should().Be(1);
    }

    // ── Write ──────────────────────────────────────────────────────────────────

    [Fact]
    public void Write_ThenReadBack_PreservesData()
    {
        var original = DataFrame.FromList(new[]
        {
            new { Name = "Alice", Age = 30 },
            new { Name = "Bob",   Age = 25 }
        });

        var path = TempFile("roundtrip.csv");
        var writer = new CsvWriter();
        writer.Write(original, path);

        var reader = new CsvReader();
        var restored = reader.Read(path);

        restored.RowCount.Should().Be(2);
        restored.ColumnCount.Should().Be(2);
        restored.GetValue(0, 0).Should().Be("Alice");
    }

    [Fact]
    public void Write_QuotesFieldsWithCommas()
    {
        var df = DataFrame.FromColumns(new()
        {
            ["Name"] = ["Smith, John"],
            ["Age"]  = [42]
        });

        var path = TempFile("commafield.csv");
        new CsvWriter().Write(df, path);

        var lines = File.ReadAllLines(path);
        lines[1].Should().StartWith("\"Smith, John\"");
    }

    // ── StreamRows ─────────────────────────────────────────────────────────────

    [Fact]
    public void StreamRows_YieldsAllDataRows()
    {
        var csv = "A,B\n1,2\n3,4\n5,6\n";
        var path = TempFile("stream.csv");
        File.WriteAllText(path, csv);

        var reader = new CsvReader();
        var rows = reader.StreamRows(path, skipHeader: true).ToList();

        rows.Should().HaveCount(3);
        rows[0].Should().Equal("1", "2");
    }

    // ── From/To stream ─────────────────────────────────────────────────────────

    [Fact]
    public void Write_ToStream_ThenRead_ProducesExpectedContent()
    {
        var df = DataFrame.FromList(new[] { new { X = 1, Y = 2 } });

        using var ms = new MemoryStream();
        new CsvWriter().Write(df, ms);
        ms.Position = 0;

        var reader = new CsvReader();
        var result = reader.Read(ms);
        result.RowCount.Should().Be(1);
    }
}
