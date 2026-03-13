using NetXLCsv.Core;
using NetXLCsv.Core.Interfaces;

namespace NetXLCsv.DataFrame;

/// <summary>
/// Static entry-point façade for creating DataFrames — mirrors pandas' top-level API.
/// All methods return <see cref="NetDataFrame"/> which is the concrete implementation
/// of <see cref="IDataFrame"/>.
/// </summary>
public static class DataFrame
{
    // ── Factory methods ───────────────────────────────────────────────────────

    /// <summary>
    /// Creates a DataFrame from a strongly-typed list.
    /// <example>
    /// <code>
    /// var df = DataFrame.FromList(new[] {
    ///     new { Name = "Alice", Age = 30 },
    ///     new { Name = "Bob",   Age = 25 }
    /// });
    /// </code>
    /// </example>
    /// </summary>
    public static NetDataFrame FromList<T>(IEnumerable<T> items) where T : class
        => NetDataFrame.FromList(items);

    /// <summary>
    /// Creates a DataFrame from a column-oriented dictionary.
    /// <example>
    /// <code>
    /// var df = DataFrame.FromColumns(new() {
    ///     ["Name"] = new object?[] { "Alice", "Bob" },
    ///     ["Age"]  = new object?[] { 30, 25 }
    /// });
    /// </code>
    /// </example>
    /// </summary>
    public static NetDataFrame FromColumns(Dictionary<string, object?[]> data)
        => NetDataFrame.FromColumns(data);

    /// <summary>
    /// Reads a CSV file and returns a DataFrame.
    /// <example>
    /// <code>
    /// var df = DataFrame.ReadCsv("sales.csv");
    /// </code>
    /// </example>
    /// </summary>
    public static NetDataFrame ReadCsv(
        string path,
        char delimiter = ',',
        bool hasHeader = true,
        bool inferTypes = true)
    {
        var reader = CsvReaderResolver.Default;
        var raw = reader.Read(path, delimiter, hasHeader);
        return (NetDataFrame)raw;
    }

    /// <summary>
    /// Reads an Excel file (.xlsx) and returns a DataFrame.
    /// <example>
    /// <code>
    /// var df = DataFrame.ReadExcel("sales.xlsx");
    /// </code>
    /// </example>
    /// </summary>
    public static NetDataFrame ReadExcel(string path, string? sheetName = null, bool hasHeader = true)
    {
        var reader = ExcelReaderResolver.Default;
        var raw = reader.ReadDataFrame(path, sheetName, hasHeader);
        return (NetDataFrame)raw;
    }

    /// <summary>Returns an empty DataFrame.</summary>
    public static NetDataFrame Empty() => NetDataFrame.Empty();
}

/// <summary>Lazy resolver for the default CSV reader.</summary>
internal static class CsvReaderResolver
{
    private static ICsvReader? _default;
    public static ICsvReader Default
    {
        get
        {
            if (_default is not null) return _default;
            var type = Type.GetType("NetXLCsv.Csv.CsvReader, NetXLCsv.Csv")
                ?? throw new InvalidOperationException("NetXLCsv.Csv assembly not loaded.");
            _default = (ICsvReader)Activator.CreateInstance(type)!;
            return _default;
        }
    }
    public static void Register(ICsvReader reader) => _default = reader;
}

/// <summary>Lazy resolver for the default Excel reader.</summary>
internal static class ExcelReaderResolver
{
    private static IExcelReader? _default;
    public static IExcelReader Default
    {
        get
        {
            if (_default is not null) return _default;
            var type = Type.GetType("NetXLCsv.Excel.ExcelReader, NetXLCsv.Excel")
                ?? throw new InvalidOperationException("NetXLCsv.Excel assembly not loaded.");
            _default = (IExcelReader)Activator.CreateInstance(type)!;
            return _default;
        }
    }
    public static void Register(IExcelReader reader) => _default = reader;
}
