using NetXLCsv.Core.Models;

namespace NetXLCsv.Core.Interfaces;

/// <summary>
/// Core contract for a tabular, in-memory data frame.
/// Implementations must be thread-safe for read-only operations.
/// </summary>
public interface IDataFrame : IEnumerable<DataRow>
{
    /// <summary>Structural description of this DataFrame.</summary>
    Schema Schema { get; }

    /// <summary>Number of data rows (excluding header).</summary>
    int RowCount { get; }

    /// <summary>Number of columns.</summary>
    int ColumnCount { get; }

    /// <summary>Returns the value at the given row/column position.</summary>
    object? GetValue(int rowIndex, int columnIndex);

    /// <summary>Returns a new DataFrame containing only the specified columns.</summary>
    IDataFrame Select(params string[] columns);

    /// <summary>Returns a new DataFrame containing only rows matching the predicate.</summary>
    IDataFrame Filter(Func<DataRow, bool> predicate);

    /// <summary>Returns a new DataFrame sorted by the given column.</summary>
    IDataFrame SortBy(string column, bool ascending = true);

    /// <summary>Returns a new DataFrame with the specified column added or replaced.</summary>
    IDataFrame AddColumn(string name, object? value);

    /// <summary>Returns a new DataFrame with the named column removed.</summary>
    IDataFrame RemoveColumn(string name);

    /// <summary>Returns groups keyed by the distinct values in <paramref name="column"/>.</summary>
    IReadOnlyDictionary<object?, IDataFrame> GroupBy(string column);

    /// <summary>Exports this DataFrame to an Excel file.</summary>
    void ToExcel(string path, string sheetName = "Sheet1");

    /// <summary>Exports this DataFrame to a CSV file.</summary>
    void ToCsv(string path, char delimiter = ',');
}
