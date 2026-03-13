using NetXLCsv.Core.Models;

namespace NetXLCsv.Core.Interfaces;

/// <summary>Represents a single worksheet inside a workbook.</summary>
public interface IWorksheet
{
    /// <summary>Name of this worksheet.</summary>
    string Name { get; }

    /// <summary>Writes a value to the given 1-based row and column.</summary>
    void WriteCell(int row, int column, object? value);

    /// <summary>Reads the raw value from the given 1-based row and column.</summary>
    object? ReadCell(int row, int column);

    /// <summary>
    /// Writes an entire table (header + data rows) starting at the given 1-based cell.
    /// </summary>
    void WriteTable(int startRow, int startColumn, IDataFrame data);

    /// <summary>Reads a rectangular range into a DataFrame.</summary>
    IDataFrame ReadTable(int startRow, int startColumn, int endRow, int endColumn);

    /// <summary>Sets the width (in character units) of the given 1-based column.</summary>
    void SetColumnWidth(int column, double width);

    /// <summary>Auto-fits the column width based on content.</summary>
    void AutoFitColumn(int column);
}
