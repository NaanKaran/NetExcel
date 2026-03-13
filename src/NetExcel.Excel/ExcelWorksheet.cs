using NetXLCsv.Core;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Models;
using NetXLCsv.Core.Utilities;
using NetXLCsv.Formatting;

namespace NetXLCsv.Excel;

/// <summary>
/// In-memory representation of a worksheet. Backed by a sparse dictionary of cells
/// so that large worksheets with few cells do not allocate unnecessary arrays.
/// </summary>
public sealed class ExcelWorksheet : IWorksheet
{
    private readonly Dictionary<(int row, int col), ExcelCell> _cells = new();
    private readonly Dictionary<int, double> _columnWidths = new();
    private readonly List<ConditionalFormat> _conditionalFormats = new();

    /// <inheritdoc/>
    public string Name { get; }

    internal ExcelWorksheet(string name) => Name = name;

    // ── IWorksheet ────────────────────────────────────────────────────────────

    /// <inheritdoc/>
    public void WriteCell(int row, int column, object? value)
    {
        ValidateCoords(row, column);
        _cells[(row, column)] = new ExcelCell(row, column, value);
    }

    /// <inheritdoc/>
    public object? ReadCell(int row, int column)
    {
        ValidateCoords(row, column);
        return _cells.TryGetValue((row, column), out var cell) ? cell.Value : null;
    }

    /// <inheritdoc/>
    public void WriteTable(int startRow, int startColumn, IDataFrame data)
    {
        Guard.NotNull(data);
        int r = startRow;

        // Header row
        for (int c = 0; c < data.ColumnCount; c++)
            WriteCell(r, startColumn + c, data.Schema.Columns[c].Name);
        r++;

        // Data rows
        foreach (var row in data)
        {
            for (int c = 0; c < data.ColumnCount; c++)
                WriteCell(r, startColumn + c, row[c]);
            r++;
        }
    }

    /// <inheritdoc/>
    public IDataFrame ReadTable(int startRow, int startColumn, int endRow, int endColumn)
    {
        int colCount = endColumn - startColumn + 1;
        var headers = new string[colCount];
        for (int c = 0; c < colCount; c++)
        {
            var h = ReadCell(startRow, startColumn + c)?.ToString();
            headers[c] = h ?? ColumnNameHelper.NumberToLetter(c + 1);
        }

        var rowData = new List<string[]>();
        for (int r = startRow + 1; r <= endRow; r++)
        {
            var vals = new string[colCount];
            for (int c = 0; c < colCount; c++)
                vals[c] = ReadCell(r, startColumn + c)?.ToString() ?? string.Empty;
            rowData.Add(vals);
        }

        return NetDataFrame.FromRawRows(headers, rowData);
    }

    /// <inheritdoc/>
    public void SetColumnWidth(int column, double width)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        _columnWidths[column] = width;
    }

    /// <inheritdoc/>
    public void AutoFitColumn(int column)
    {
        // Heuristic: find max string length in the column and estimate width
        int maxLen = 0;
        foreach (var ((r, c), cell) in _cells)
        {
            if (c != column) continue;
            var len = cell.Value?.ToString()?.Length ?? 0;
            if (len > maxLen) maxLen = len;
        }
        _columnWidths[column] = Math.Max(8, maxLen * 1.2);
    }

    // ── Styling shortcuts ─────────────────────────────────────────────────────

    /// <summary>Applies a style to a specific cell.</summary>
    public void SetCellStyle(int row, int column, CellStyle style)
    {
        if (!_cells.TryGetValue((row, column), out var cell))
        {
            cell = new ExcelCell(row, column);
            _cells[(row, column)] = cell;
        }
        cell.Style = style;
    }

    /// <summary>Applies header style to the first row starting at the given column.</summary>
    public void SetHeaderStyle(bool bold = true, string backgroundColor = "#EEEEEE")
    {
        var style = new CellStyle
        {
            Font = bold ? FontStyle.Bold : FontStyle.Default,
            BackgroundColor = backgroundColor,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        int maxCol = _cells.Keys.Max(k => k.col);
        for (int c = 1; c <= maxCol; c++)
            SetCellStyle(1, c, style);
    }

    /// <summary>Adds a conditional format rule to this worksheet.</summary>
    public void AddConditionalFormat(ConditionalFormat format) =>
        _conditionalFormats.Add(format);

    // ── Internal accessors (for ExcelWorkbook serialization) ──────────────────

    internal IReadOnlyDictionary<(int row, int col), ExcelCell> Cells => _cells;
    internal IReadOnlyDictionary<int, double> ColumnWidths => _columnWidths;
    internal IReadOnlyList<ConditionalFormat> ConditionalFormats => _conditionalFormats;

    internal int MaxRow => _cells.Count == 0 ? 0 : _cells.Keys.Max(k => k.row);
    internal int MaxColumn => _cells.Count == 0 ? 0 : _cells.Keys.Max(k => k.col);

    private static void ValidateCoords(int row, int col)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row), "Row must be >= 1.");
        if (col < 1) throw new ArgumentOutOfRangeException(nameof(col), "Column must be >= 1.");
    }
}
