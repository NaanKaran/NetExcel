using NetXLCsv.Core.Models;
using NetXLCsv.Formatting;

namespace NetXLCsv.Excel;

/// <summary>
/// Represents a mutable cell in an <see cref="ExcelWorksheet"/>.
/// </summary>
public sealed class ExcelCell
{
    private object? _value;

    /// <summary>1-based row position.</summary>
    public int Row { get; }

    /// <summary>1-based column position.</summary>
    public int Column { get; }

    /// <summary>The cell value. Setting this marks the parent worksheet as dirty.</summary>
    public object? Value
    {
        get => _value;
        set => _value = value;
    }

    /// <summary>Optional style override for this cell.</summary>
    public CellStyle? Style { get; set; }

    /// <summary>Optional formula (without leading '=').</summary>
    public string? Formula { get; set; }

    /// <summary>A1 address of this cell (e.g. "A1", "B12").</summary>
    public string Address => new CellAddress(Row - 1, Column - 1).ToA1Notation();

    public ExcelCell(int row, int column, object? value = null)
    {
        Row = row;
        Column = column;
        _value = value;
    }

    /// <inheritdoc/>
    public override string ToString() => $"{Address}={Value ?? "<null>"}";
}
