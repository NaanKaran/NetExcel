using BenchmarkDotNet.Attributes;
using NetXLCsv.Streaming;
using OfficeOpenXml;

namespace NetXLCsv.Benchmarks;

/// <summary>
/// Compares NetXLCsv streaming writer vs EPPlus for bulk Excel writes.
/// Run in Release mode: dotnet run -c Release --project benchmarks/NetExcel.Benchmarks
///
/// NOTE: ClosedXML is intentionally excluded.
/// ClosedXML 0.102.x requires DocumentFormat.OpenXml >= 2.16 &amp;&amp; &lt; 3.0,
/// which conflicts with NetXLCsv's dependency on DocumentFormat.OpenXml 3.1.0 (NU1107).
/// To benchmark against ClosedXML, create a separate isolated project without NetExcel
/// project references. Reference code is in the region comment at the bottom of this file.
/// </summary>
[SimpleJob] // targets the host runtime automatically (net10.0)
[MemoryDiagnoser]
[MinColumn, MaxColumn, MeanColumn, MedianColumn]
public class ExcelWriteBenchmarks
{
    [Params(100_000, 1_000_000)]
    public int RowCount { get; set; }

    private static readonly string[] Headers = ["Id", "Name", "Revenue", "Country", "Active"];
    private static readonly object?[] DataRow = [1, "Alice Johnson", 99_999.99, "India", true];

    private string _outPath = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        _outPath = Path.Combine(Path.GetTempPath(), $"bench_{Guid.NewGuid():N}.xlsx");
        // EPPlus requires license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_outPath)) File.Delete(_outPath);
    }

    // ── NetExcel Streaming ────────────────────────────────────────────────────

    [Benchmark(Baseline = true)]
    public void NetXLCsv_StreamingWriter()
    {
        using var writer = StreamingExcelWriter.Create(_outPath);
        writer.WriteHeader(Headers);
        for (int i = 0; i < RowCount; i++)
            writer.WriteRow(DataRow);
    }

    // ── EPPlus ────────────────────────────────────────────────────────────────

    [Benchmark]
    public void EPPlus_Write()
    {
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Data");

        for (int c = 0; c < Headers.Length; c++)
            ws.Cells[1, c + 1].Value = Headers[c];

        for (int r = 0; r < RowCount; r++)
        {
            for (int c = 0; c < DataRow.Length; c++)
                ws.Cells[r + 2, c + 1].Value = DataRow[c];
        }

        pkg.SaveAs(new FileInfo(_outPath));
    }

    /*
     * ── ClosedXML reference implementation (cannot be compiled here) ───────────
     *
     * Reason: ClosedXML 0.102.x pins DocumentFormat.OpenXml >= 2.16.0 && < 3.0.0.
     * NetXLCsv uses DocumentFormat.OpenXml 3.1.0. These constraints are mutually
     * exclusive — NuGet raises NU1107 and restore fails.
     *
     * To run a ClosedXML benchmark, create a standalone project with no NetExcel
     * project references, add only ClosedXML, and use this code:
     *
     *   using ClosedXML.Excel;
     *
     *   using var wb = new XLWorkbook();
     *   var ws = wb.AddWorksheet("Data");
     *   for (int c = 0; c < headers.Length; c++)
     *       ws.Cell(1, c + 1).Value = headers[c];
     *   for (int r = 0; r < rowCount; r++)
     *       for (int c = 0; c < data.Length; c++)
     *           ws.Cell(r + 2, c + 1).Value = data[c]?.ToString();
     *   wb.SaveAs(outputPath);
     */
}
