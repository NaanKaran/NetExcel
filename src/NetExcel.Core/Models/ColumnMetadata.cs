namespace NetXLCsv.Core.Models;

/// <summary>
/// Describes a single column in a DataFrame or worksheet.
/// </summary>
public sealed class ColumnMetadata
{
    /// <summary>Column name / header.</summary>
    public string Name { get; init; }

    /// <summary>Zero-based column index.</summary>
    public int Index { get; init; }

    /// <summary>Inferred or declared data type of this column.</summary>
    public DataType DataType { get; init; }

    /// <summary>Whether null values are permitted.</summary>
    public bool IsNullable { get; init; } = true;

    /// <summary>Optional display format string (e.g. "yyyy-MM-dd", "0.00").</summary>
    public string? Format { get; init; }

    public ColumnMetadata(string name, int index, DataType dataType = DataType.Unknown)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Name = name;
        Index = index;
        DataType = dataType;
    }

    /// <inheritdoc/>
    public override string ToString() => $"[{Index}] {Name} ({DataType})";
}
