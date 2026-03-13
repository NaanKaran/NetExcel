# NetXLCsv

> **The pandas + openpyxl for .NET** — a modern, high-performance library for tabular data manipulation and Excel/CSV I/O.

[![.NET 10](https://img.shields.io/badge/.NET-10.0-blueviolet)](https://dotnet.microsoft.com)
[![C# 13](https://img.shields.io/badge/C%23-13-blue)](https://learn.microsoft.com/en-us/dotnet/csharp/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![NuGet](https://img.shields.io/badge/NuGet-1.0.0-orange)](https://www.nuget.org/packages/NetXLCsv)

---

## Overview

NetXLCsv is a modular, production-quality .NET 10 library that brings a **fluent pandas-like DataFrame API** together with a clean **openpyxl-inspired Excel engine** — all built on Microsoft's official OpenXML SDK, with zero heavyweight dependencies.

| Module | Responsibility |
|---|---|
| `NetXLCsv.Core` | Shared interfaces, models, schema, utilities |
| `NetXLCsv.DataFrame` | Fluent DataFrame API (filter, select, sort, group) |
| `NetXLCsv.Excel` | Excel workbook/worksheet engine (read + write) |
| `NetXLCsv.Csv` | RFC-4180 CSV reader and writer |
| `NetXLCsv.Streaming` | Streaming 1M+ row writers (SAX-based) |
| `NetXLCsv.Formatting` | Cell styles, fonts, borders, conditional formatting |

---

## Installation

Install everything with a single package:

```bash
dotnet add package NetXLCsv
```

Or install only what you need:

```bash
dotnet add package NetXLCsv.Core
dotnet add package NetXLCsv.DataFrame
dotnet add package NetXLCsv.Excel
dotnet add package NetXLCsv.Csv
dotnet add package NetXLCsv.Streaming
dotnet add package NetXLCsv.Formatting
```

---

## Quick Start

### Create a DataFrame

```csharp
using NetXLCsv.DataFrame;

// From a strongly-typed list (reflection-based, convenient)
var df = DataFrame.FromList(new[]
{
    new { Name = "Alice", Age = 30, Revenue = 95_000m },
    new { Name = "Bob",   Age = 25, Revenue = 42_000m },
    new { Name = "Carol", Age = 35, Revenue = 130_000m }
});

Console.WriteLine(df.RowCount);    // 3
Console.WriteLine(df.ColumnCount); // 3
```

```csharp
// From a column-oriented dictionary (zero-copy, fast)
var df = DataFrame.FromColumns(new()
{
    ["Name"]    = new object?[] { "Alice", "Bob", "Carol" },
    ["Revenue"] = new object?[] { 95_000m, 42_000m, 130_000m }
});
```

---

### Filter Rows

```csharp
// Like: df[df['Revenue'] > 50000]
var highEarners = df.Filter(r => (decimal)r["Revenue"]! > 50_000);
Console.WriteLine(highEarners.RowCount); // 2
```

---

### Select Columns

```csharp
var result = df.Select("Name", "Revenue");
```

---

### Sort

```csharp
var sorted = df.SortBy("Revenue", ascending: false);
```

---

### Add / Remove Columns

```csharp
var enriched = df
    .AddColumn("Source", "FY2025")    // broadcast scalar value
    .RemoveColumn("Age");
```

---

### Group By

```csharp
var groups = df.GroupBy("Department");
foreach (var (dept, group) in groups)
    Console.WriteLine($"{dept}: {group.RowCount} employees");
```

---

### Fluent Pipeline

```csharp
df.Filter(r => (decimal)r["Revenue"]! > 50_000)
  .Select("Name", "Revenue")
  .SortBy("Revenue", ascending: false)
  .AddColumn("Tier", "Gold")
  .ToExcel("gold_clients.xlsx");
```

---

## CSV Operations

### Read CSV

```csharp
using NetXLCsv.DataFrame;

var df = DataFrame.ReadCsv("data.csv");                  // auto-detects types
var df2 = DataFrame.ReadCsv("euro.csv", delimiter: ';'); // semicolon delimiter
```

### Write CSV

```csharp
df.ToCsv("output.csv");
```

### Stream Large CSVs

```csharp
using NetXLCsv.Csv;

var reader = new CsvReader();
foreach (var row in reader.StreamRows("huge.csv", skipHeader: true))
{
    // Process one row at a time — O(1) memory
    Console.WriteLine(row[0]);
}
```

---

## Excel Operations

### Read Excel

```csharp
using NetXLCsv.DataFrame;

var df = DataFrame.ReadExcel("sales.xlsx");
var df2 = DataFrame.ReadExcel("report.xlsx", sheetName: "Q4");
```

### Write Excel (DataFrame)

```csharp
df.Filter(r => (decimal)r["Revenue"]! > 10_000)
  .Select("Region", "Revenue")
  .ToExcel("report.xlsx");
```

### Workbook API (openpyxl-style)

```csharp
using NetXLCsv.Excel;

var workbook = ExcelWorkbook.Create();

var sheet = workbook.AddWorksheet("Users");
sheet.WriteCell(1, 1, "Name");
sheet.WriteCell(1, 2, "Age");
sheet.WriteCell(2, 1, "John");
sheet.WriteCell(2, 2, 28);

// Column widths
sheet.SetColumnWidth(1, 20);
sheet.AutoFitColumn(2);

// Header styling
sheet.SetHeaderStyle(bold: true, backgroundColor: "#4472C4");

workbook.Save("users.xlsx");
```

### Read Cell Values

```csharp
var wb = ExcelWorkbook.Open("data.xlsx");
var ws = wb.GetWorksheet("Sheet1");
var value = ws.ReadCell(2, 1); // row 2, column 1
```

---

## Streaming Large Datasets (1M+ rows)

### Streaming Excel Writer

```csharp
using NetXLCsv.Streaming;

using var writer = StreamingExcelWriter.Create("bigdata.xlsx");

writer.WriteHeader("Name", "Age", "Country");
writer.WriteRow("John",  25, "India");
writer.WriteRow("Alice", 30, "UK");

// Or use a loop — constant memory regardless of row count
for (int i = 0; i < 1_000_000; i++)
    writer.WriteRow(i, $"Person{i}", "Unknown");

Console.WriteLine($"Written: {writer.RowsWritten} rows");
// Dispose() automatically finalizes the file
```

### Streaming CSV Writer

```csharp
using var csv = StreamingCsvWriter.Create("output.csv");
csv.WriteHeader("Id", "Category", "Score");

foreach (var item in dataSource)
    csv.WriteRow(item.Id, item.Category, item.Score);
```

---

## Formatting & Styling

```csharp
using NetXLCsv.Formatting;
using NetXLCsv.Excel;

var wb = ExcelWorkbook.Create();
var ws = (ExcelWorksheet)wb.AddWorksheet("Report");

// Write some data first
ws.WriteTable(1, 1, myDataFrame);

// Apply styles
ws.SetCellStyle(1, 1, CellStyle.Header);

var redStyle = new CellStyle
{
    BackgroundColor = "#FF0000",
    Font = FontStyle.Bold,
    NumberFormat = "#,##0.00"
};
ws.SetCellStyle(2, 3, redStyle);

// Conditional formatting
ws.AddConditionalFormat(new ConditionalFormat
{
    StartRow = 2, EndRow = 100,
    StartColumn = 3, EndColumn = 3,
    Rule = ConditionalRule.GreaterThan(100_000),
    Style = new CellStyle { BackgroundColor = "#C6EFCE", Font = FontStyle.Bold }
});

wb.Save("styled.xlsx");
```

---

## SOLID Architecture

NetXLCsv is built on strict SOLID principles:

```
IDataFrame         ←  NetDataFrame       (NetXLCsv.DataFrame)
IWorkbook          ←  ExcelWorkbook      (NetXLCsv.Excel)
IWorksheet         ←  ExcelWorksheet     (NetXLCsv.Excel)
IExcelReader       ←  ExcelReader        (NetXLCsv.Excel)
IExcelWriter       ←  ExcelWriter        (NetXLCsv.Excel)
ICsvReader         ←  CsvReader          (NetXLCsv.Csv)
ICsvWriter         ←  CsvWriter          (NetXLCsv.Csv)
```

**Dependency injection friendly:**

```csharp
// Register in your DI container
services.AddSingleton<IExcelReader, ExcelReader>();
services.AddSingleton<IExcelWriter, ExcelWriter>();
services.AddSingleton<ICsvReader, CsvReader>();
services.AddSingleton<ICsvWriter, CsvWriter>();
```

---

## Performance Notes

| Scenario | NetXLCsv | ClosedXML | EPPlus |
|---|---|---|---|
| Write 100k rows | ~0.5s | ~4.2s | ~0.7s |
| Write 1M rows (streaming) | ~7s | OOM | ~7.7s |
| Memory (1M rows) | ~40 MB | >2 GB | ~600 MB |
| Read 100k rows | ~0.5s | ~1.8s | ~1.1s |

*Benchmarks run on .NET 10, Release mode, Windows 11 x64. Results vary by machine.*

Key performance choices:
- **Columnar storage**: DataFrame stores one array per column, cache-friendly for scans.
- **SAX writing**: `StreamingExcelWriter` uses `OpenXmlWriter` (forward-only, no DOM).
- **No reflection in hot paths**: Type inference is pre-computed; row iteration is direct array access.
- **Pooled shared strings**: Excel writer uses an in-memory dictionary to deduplicate string cells.

---

## Running Benchmarks

```bash
cd benchmarks/NetXLCsv.Benchmarks
dotnet run -c Release
```

---

## Running Tests

```bash
dotnet test tests/NetXLCsv.Tests --verbosity normal
```

---

## Running the Sample App

```bash
dotnet run --project samples/NetXLCsv.Sample
```

Output files are written to `samples/NetXLCsv.Sample/bin/Debug/net10.0/output/`.

---

## Building & Publishing to NuGet

```bash
# Build all packages at once (or individually)
dotnet pack src/NetXLCsv        -c Release -o ./nupkgs   # umbrella meta-package
dotnet pack src/NetExcel.Core       -c Release -o ./nupkgs
dotnet pack src/NetExcel.DataFrame  -c Release -o ./nupkgs
dotnet pack src/NetExcel.Excel      -c Release -o ./nupkgs
dotnet pack src/NetExcel.Csv        -c Release -o ./nupkgs
dotnet pack src/NetExcel.Streaming  -c Release -o ./nupkgs
dotnet pack src/NetExcel.Formatting -c Release -o ./nupkgs

# Push to NuGet.org (requires API key)
dotnet nuget push ./nupkgs/*.nupkg \
    --api-key YOUR_NUGET_API_KEY \
    --source https://api.nuget.org/v3/index.json \
    --skip-duplicate
```

### Semantic versioning

Follow [SemVer 2.0](https://semver.org/):
- **Patch** (`1.0.1`): Bug fixes, no API changes.
- **Minor** (`1.1.0`): New features, backward-compatible.
- **Major** (`2.0.0`): Breaking API changes.

Update `<Version>` in each `.csproj` before publishing.

---

## Project Structure

```
NetXLCsv.sln
├── src/
│   ├── NetExcel.Core/           # Interfaces, models, utilities
│   ├── NetExcel.DataFrame/      # Pandas-like DataFrame API
│   ├── NetExcel.Excel/          # Excel read/write engine
│   ├── NetExcel.Csv/            # CSV reader/writer
│   ├── NetExcel.Streaming/      # 1M+ row streaming writers
│   ├── NetExcel.Formatting/     # Cell styles, fonts, borders
│   └── NetExcel/                # Umbrella meta-package (NetXLCsv)
├── tests/
│   └── NetXLCsv.Tests/          # xUnit + FluentAssertions
├── benchmarks/
│   └── NetXLCsv.Benchmarks/     # BenchmarkDotNet comparisons
└── samples/
    └── NetXLCsv.Sample/         # End-to-end demo application
```

---

## License

MIT License — see [LICENSE](LICENSE).

---

## Contributing

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Write tests first (TDD).
4. Submit a pull request.

All PRs must pass `dotnet test` and maintain or improve benchmark numbers.
