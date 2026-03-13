// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  NetXLCsv Sample Application                                            ║
// ║  Demonstrates the full pandas-like workflow for .NET                    ║
// ╚══════════════════════════════════════════════════════════════════════════╝

using NetXLCsv.DataFrame;
using NetXLCsv.Excel;
using NetXLCsv.Streaming;

var outputDir = Path.Combine(AppContext.BaseDirectory, "output");
Directory.CreateDirectory(outputDir);

Console.WriteLine("╔══════════════════════════════════════╗");
Console.WriteLine("║        NetXLCsv Demo Runner          ║");
Console.WriteLine("╚══════════════════════════════════════╝\n");

// ─────────────────────────────────────────────────────────────────────────────
// 1. CREATE A DATAFRAME FROM A LIST (like pandas.DataFrame.from_records)
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("▶ 1. Creating DataFrame from list...");

var salesData = Enumerable.Range(1, 30).Select(i => new
{
    Id = i,
    Region = i % 3 == 0 ? "APAC" : i % 3 == 1 ? "EMEA" : "AMER",
    Product = $"Product{i % 5 + 1}",
    Revenue = Math.Round((decimal)(i * 5432.10 + (i % 7) * 1234), 2),
    Active = i % 4 != 0
}).ToList();

var df = DataFrame.FromList(salesData);
Console.WriteLine($"   Shape: {df.RowCount} rows × {df.ColumnCount} cols");
Console.WriteLine($"   Columns: {string.Join(", ", df.Schema.Columns.Select(c => c.Name))}");

// ─────────────────────────────────────────────────────────────────────────────
// 2. FILTER  (like df[df['Revenue'] > 50000])
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 2. Filtering rows where Revenue > 50,000...");
var highRevenue = df.Filter(r => (decimal)r["Revenue"]! > 50_000);
Console.WriteLine($"   Matched: {highRevenue.RowCount} rows");

// ─────────────────────────────────────────────────────────────────────────────
// 3. SELECT COLUMNS  (like df[['Region','Revenue']])
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 3. Selecting columns: Region, Revenue...");
var subset = highRevenue.Select("Region", "Revenue");
Console.WriteLine($"   Columns now: {string.Join(", ", subset.Schema.Columns.Select(c => c.Name))}");

// ─────────────────────────────────────────────────────────────────────────────
// 4. ADD A COLUMN  (like df.assign(Source='FY2025'))
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 4. Adding column 'Source' = 'FY2025'...");
var enriched = subset.AddColumn("Source", "FY2025");
Console.WriteLine($"   New column count: {enriched.ColumnCount}");

// ─────────────────────────────────────────────────────────────────────────────
// 5. SORT  (like df.sort_values('Revenue', ascending=False))
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 5. Sorting by Revenue descending...");
var sorted = enriched.SortBy("Revenue", ascending: false);
Console.WriteLine($"   Top Revenue: {sorted.GetValue(0, 1)}");

// ─────────────────────────────────────────────────────────────────────────────
// 6. GROUP BY  (like df.groupby('Region').size())
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 6. GroupBy Region...");
var groups = df.GroupBy("Region");
foreach (var (region, group) in groups)
    Console.WriteLine($"   {region}: {group.RowCount} rows");

// ─────────────────────────────────────────────────────────────────────────────
// 7. EXPORT TO EXCEL  (like df.to_excel('report.xlsx'))
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 7. Exporting filtered data to Excel...");
var excelPath = Path.Combine(outputDir, "sales_report.xlsx");
sorted.ToExcel(excelPath);
Console.WriteLine($"   Saved → {excelPath}");

// ─────────────────────────────────────────────────────────────────────────────
// 8. EXPORT TO CSV  (like df.to_csv('data.csv'))
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 8. Exporting full dataset to CSV...");
var csvPath = Path.Combine(outputDir, "sales_full.csv");
df.ToCsv(csvPath);
Console.WriteLine($"   Saved → {csvPath}");

// ─────────────────────────────────────────────────────────────────────────────
// 9. READ BACK EXCEL  (like pd.read_excel('report.xlsx'))
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 9. Reading back the Excel file...");
var restored = DataFrame.ReadExcel(excelPath);
Console.WriteLine($"   Read: {restored.RowCount} rows × {restored.ColumnCount} cols");

// ─────────────────────────────────────────────────────────────────────────────
// 10. WORKBOOK API  (openpyxl-style)
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 10. Workbook API (openpyxl-style)...");
var wbPath = Path.Combine(outputDir, "manual_workbook.xlsx");
using (var workbook = ExcelWorkbook.Create())
{
    var sheet = (ExcelWorksheet)workbook.AddWorksheet("Users");

    // Write headers
    sheet.WriteCell(1, 1, "Name");
    sheet.WriteCell(1, 2, "Age");
    sheet.WriteCell(1, 3, "Score");

    // Write data
    sheet.WriteCell(2, 1, "John");  sheet.WriteCell(2, 2, 28); sheet.WriteCell(2, 3, 95.5);
    sheet.WriteCell(3, 1, "Alice"); sheet.WriteCell(3, 2, 35); sheet.WriteCell(3, 3, 88.2);
    sheet.WriteCell(4, 1, "Bob");   sheet.WriteCell(4, 2, 22); sheet.WriteCell(4, 3, 76.9);

    // Styling
    sheet.SetColumnWidth(1, 20);
    sheet.SetColumnWidth(2, 8);
    sheet.SetColumnWidth(3, 10);
    sheet.SetHeaderStyle(bold: true, backgroundColor: "#4472C4");

    workbook.Save(wbPath);
}
Console.WriteLine($"   Saved → {wbPath}");

// ─────────────────────────────────────────────────────────────────────────────
// 11. STREAMING EXCEL (1M+ rows demo with 10k rows)
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 11. Streaming Excel Writer (10,000 rows)...");
var streamPath = Path.Combine(outputDir, "streaming_10k.xlsx");
var sw = System.Diagnostics.Stopwatch.StartNew();

using (var writer = StreamingExcelWriter.Create(streamPath))
{
    writer.WriteHeader("Id", "Name", "Revenue", "Country", "Date");
    var rnd = new Random(42);
    var countries = new[] { "India", "UK", "USA", "Germany", "Japan" };
    for (int i = 1; i <= 10_000; i++)
    {
        writer.WriteRow(
            i,
            $"Customer {i}",
            Math.Round(rnd.NextDouble() * 200_000, 2),
            countries[rnd.Next(countries.Length)],
            DateTime.Today.AddDays(-rnd.Next(365))
        );
    }
}

sw.Stop();
var fi = new FileInfo(streamPath);
Console.WriteLine($"   10k rows in {sw.ElapsedMilliseconds}ms → {fi.Length / 1024} KB");
Console.WriteLine($"   Saved → {streamPath}");

// ─────────────────────────────────────────────────────────────────────────────
// 12. STREAMING CSV
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n▶ 12. Streaming CSV Writer (100,000 rows)...");
var bigCsvPath = Path.Combine(outputDir, "streaming_100k.csv");
sw.Restart();

using (var csvWriter = StreamingCsvWriter.Create(bigCsvPath))
{
    csvWriter.WriteHeader("Id", "Category", "Value", "Timestamp");
    for (int i = 1; i <= 100_000; i++)
        csvWriter.WriteRow(i, $"Cat{i % 20}", i * 0.789, DateTime.UtcNow.AddSeconds(-i));
}

sw.Stop();
var csvFi = new FileInfo(bigCsvPath);
Console.WriteLine($"   100k rows in {sw.ElapsedMilliseconds}ms → {csvFi.Length / 1024} KB");

// ─────────────────────────────────────────────────────────────────────────────
// SUMMARY
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n✅ All demos completed. Output files:");
foreach (var f in Directory.GetFiles(outputDir))
    Console.WriteLine($"   {Path.GetFileName(f),30}  ({new FileInfo(f).Length / 1024,5} KB)");
