namespace NetXLCsv.Core.Models;

/// <summary>
/// Describes the structure (column names + types) of a DataFrame or worksheet.
/// Immutable after construction.
/// </summary>
public sealed class Schema
{
    private readonly ColumnMetadata[] _columns;
    private readonly Dictionary<string, int> _nameIndex;

    public Schema(IEnumerable<ColumnMetadata> columns)
    {
        _columns = columns.ToArray();
        _nameIndex = new Dictionary<string, int>(_columns.Length, StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < _columns.Length; i++)
            _nameIndex[_columns[i].Name] = i;
    }

    /// <summary>All column descriptors in order.</summary>
    public IReadOnlyList<ColumnMetadata> Columns => _columns;

    /// <summary>Number of columns.</summary>
    public int ColumnCount => _columns.Length;

    /// <summary>Returns the column descriptor for the given name (case-insensitive).</summary>
    public ColumnMetadata GetColumn(string name)
    {
        if (_nameIndex.TryGetValue(name, out var idx))
            return _columns[idx];
        throw new KeyNotFoundException($"Column '{name}' not found in schema.");
    }

    /// <summary>Returns true if the schema contains a column with the given name.</summary>
    public bool HasColumn(string name) => _nameIndex.ContainsKey(name);

    /// <summary>Returns the zero-based index of the named column.</summary>
    public int IndexOf(string name)
    {
        if (_nameIndex.TryGetValue(name, out var idx)) return idx;
        throw new KeyNotFoundException($"Column '{name}' not found.");
    }

    /// <summary>Returns a sub-schema containing only the requested column names.</summary>
    public Schema Select(IEnumerable<string> names)
    {
        var cols = names.Select((n, i) =>
        {
            var c = GetColumn(n);
            return new ColumnMetadata(c.Name, i, c.DataType) { IsNullable = c.IsNullable, Format = c.Format };
        });
        return new Schema(cols);
    }

    /// <summary>Returns a copy of the internal name→index dictionary (used by DataRow).</summary>
    public IReadOnlyDictionary<string, int> NameIndex => _nameIndex;

    /// <inheritdoc/>
    public override string ToString() =>
        "Schema(" + string.Join(", ", _columns.Select(c => c.ToString())) + ")";
}
