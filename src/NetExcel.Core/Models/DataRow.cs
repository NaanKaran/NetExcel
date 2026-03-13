using System.Runtime.CompilerServices;

namespace NetXLCsv.Core.Models;

/// <summary>
/// A single row in a DataFrame — a thin, allocation-conscious wrapper over a value array.
/// Column access is O(1) via the pre-built name→index map.
/// </summary>
public sealed class DataRow
{
    private readonly object?[] _values;
    private readonly IReadOnlyDictionary<string, int> _columnIndex;

    public DataRow(object?[] values, IReadOnlyDictionary<string, int> columnIndex)
    {
        _values = values;
        _columnIndex = columnIndex;
    }

    /// <summary>Gets a value by column name.</summary>
    public object? this[string column]
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            if (!_columnIndex.TryGetValue(column, out var idx))
                throw new KeyNotFoundException($"Column '{column}' does not exist.");
            return _values[idx];
        }
    }

    /// <summary>Gets a value by zero-based column index.</summary>
    public object? this[int index]
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _values[index];
    }

    /// <summary>Returns the typed value for a column, or the default if null.</summary>
    public T? Get<T>(string column) => (T?)this[column];

    /// <summary>Number of values in this row.</summary>
    public int Count => _values.Length;

    /// <summary>Raw values array (treat as read-only).</summary>
    public ReadOnlySpan<object?> Values => _values;

    /// <summary>Returns true if all values are null or empty strings.</summary>
    public bool IsEmpty()
    {
        foreach (var v in _values)
        {
            if (v is not null && v is not string s || (v is string str && str.Length > 0))
                return false;
        }
        return true;
    }

    /// <inheritdoc/>
    public override string ToString()
    {
        var parts = new string[_values.Length];
        foreach (var (name, idx) in _columnIndex)
            parts[idx] = $"{name}={_values[idx] ?? "null"}";
        return "{ " + string.Join(", ", parts) + " }";
    }
}
