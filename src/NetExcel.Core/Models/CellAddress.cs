namespace NetXLCsv.Core.Models;

/// <summary>
/// Represents a zero-based (Row, Column) cell address.
/// </summary>
public readonly record struct CellAddress(int Row, int Column)
{
    /// <summary>Converts to Excel A1 notation (e.g. row=0,col=0 → "A1").</summary>
    public string ToA1Notation()
    {
        var col = Column;
        var colLabel = string.Empty;
        do
        {
            colLabel = (char)('A' + col % 26) + colLabel;
            col = col / 26 - 1;
        }
        while (col >= 0);
        return colLabel + (Row + 1);
    }

    /// <inheritdoc/>
    public override string ToString() => ToA1Notation();
}
