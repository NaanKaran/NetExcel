# NetXLCsv — Benchmarks

**NetXLCsv** is a high-performance .NET library for generating Excel (`.xlsx`) and CSV files with a focus on **streaming writes, low overhead, and scalability for large datasets**.

It is designed as a lightweight alternative to traditional Excel libraries while supporting **large dataset exports efficiently**.

---

# ✨ Features

* 🚀 High-performance Excel and CSV generation
* 📦 Lightweight and dependency-minimal
* 🧵 Streaming writer for large datasets (1M+ rows)
* ⚡ Optimized for large row counts
* 🧪 Fully benchmarked using BenchmarkDotNet
* 🧩 Designed for .NET 10+

---

# 📊 Performance Benchmarks

Benchmarks were executed using **BenchmarkDotNet** comparing **NetXLCsv** against **EPPlus** for Excel file generation.

Environment:

* **OS:** Windows 11
* **Runtime:** .NET 10.0.3
* **Architecture:** x64 RyuJIT AVX2
* **Benchmark Tool:** BenchmarkDotNet v0.14.0

Test scenario:

* Generate Excel file
* Sequential row writes
* Streaming writer implementation
* Two dataset sizes:

  * 100,000 rows
  * 1,000,000 rows

---

# ⚡ Benchmark Results

| Method                    | Rows | Mean Time  | Memory Allocated |
| ------------------------- | ---- | ---------- | ---------------- |
| NetXLCsv StreamingWriter  | 100K | **471 ms** | 396 MB           |
| EPPlus                    | 100K | 663 ms     | 333 MB           |
| NetXLCsv StreamingWriter  | 1M   | **6.99 s** | 3814 MB          |
| EPPlus                    | 1M   | 7.65 s     | 3139 MB          |

---

# 🚀 Performance Comparison

## 100K Rows

| Library       | Time       |
| ------------- | ---------- |
| **NetXLCsv**  | **471 ms** |
| EPPlus        | 663 ms     |

NetXLCsv is approximately **29% faster** than EPPlus when generating Excel files with 100,000 rows.

---

## 1 Million Rows

| Library       | Time       |
| ------------- | ---------- |
| **NetXLCsv**  | **6.99 s** |
| EPPlus        | 7.65 s     |

NetXLCsv is approximately **9% faster** than EPPlus for large datasets.

---

# 📈 Summary

NetXLCsv demonstrates strong performance improvements compared to EPPlus:

* ⚡ **29% faster** for medium datasets (100K rows)
* ⚡ **~9% faster** for large datasets (1M rows)
* 🚀 Designed for scalable Excel generation
* 🧵 Optimized streaming architecture

---

# 🧪 Benchmark Methodology

Benchmarks were implemented using `BenchmarkDotNet`.

Key measurement metrics:

* Execution time
* Memory allocation
* Garbage collection pressure
* CPU stability

Benchmark parameters:

```
RowCount = 100000
RowCount = 1000000
```

Each benchmark executed multiple iterations to ensure statistically reliable results.

---

# 📂 Benchmark Reports

Detailed benchmark reports are available in:

```
BenchmarkDotNet.Artifacts/results/
```

Generated files:

* `NetXLCsv.Benchmarks.ExcelWriteBenchmarks-report.html`
* `NetXLCsv.Benchmarks.ExcelWriteBenchmarks-report.csv`
* `NetXLCsv.Benchmarks.ExcelWriteBenchmarks-report-github.md`

---

# 📦 Installation

```bash
dotnet add package NetXLCsv
```

---

# 🧑‍💻 Example Usage

```csharp
using NetXLCsv.Streaming;

using var writer = StreamingExcelWriter.Create("output.xlsx");

writer.WriteHeader("Id", "Name", "Timestamp");

for (int i = 0; i < 100_000; i++)
{
    writer.WriteRow(i, $"User {i}", DateTime.UtcNow);
}

Console.WriteLine($"Written: {writer.RowsWritten} rows");
```

---

# 🚀 Running Benchmarks

```bash
cd benchmarks/NetXLCsv.Benchmarks
dotnet run -c Release
```

---

# 🎯 Roadmap

Planned improvements:

* Async streaming support
* Lower memory allocation
* Additional conditional formatting rules
* Chart generation API
* Additional performance optimizations

---

# 📜 License

MIT License

---

# 🤝 Contributing

Contributions are welcome.
Please open issues or pull requests for improvements and feature suggestions.

---
