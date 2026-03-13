using System.Collections;
using System.Reflection;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Models;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Core;

/// <summary>
/// Columnar, immutable-by-default DataFrame.
/// Lives in NetXLCsv.Core so that NetXLCsv.Csv and NetXLCsv.Excel can create
/// DataFrame instances without a circular project dependency on NetXLCsv.DataFrame.
///
/// All transformation methods (Filter, Select, AddColumn, etc.) return new instances —
/// the original is never mutated, enabling safe functional pipelines.
///
/// Internal storage: one object?[] per column — cache-friendly for column scans.
/// </summary>
public sealed class NetDataFrame : IDataFrame
{
    // ── Columnar storage ──────────────────────────────────────────────────────
    private readonly string[] _columnNames;
    private readonly object?[][] _columns;   // _columns[colIdx][rowIdx]
    private readonly Schema _schema;
    private readonly Dictionary<string, int> _nameIndex;

    // ── Construction ──────────────────────────────────────────────────────────

    private NetDataFrame(string[] columnNames, object?[][] columns, Schema schema)
    {
        _columnNames = columnNames;
        _columns = columns;
        _schema = schema;
        _nameIndex = new Dictionary<string, int>(columnNames.Length, StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < columnNames.Length; i++)
            _nameIndex[columnNames[i]] = i;
    }

    // ── IDataFrame ────────────────────────────────────────────────────────────

    public Schema Schema => _schema;
    public int RowCount => _columns.Length == 0 ? 0 : _columns[0].Length;
    public int ColumnCount => _columnNames.Length;

    public object? GetValue(int rowIndex, int columnIndex)
    {
        Guard.IndexInRange(rowIndex, RowCount);
        Guard.IndexInRange(columnIndex, ColumnCount);
        return _columns[columnIndex][rowIndex];
    }

    public IDataFrame Select(params string[] columns)
    {
        Guard.NotNull(columns);
        if (columns.Length == 0) return this;

        var newCols = new object?[columns.Length][];
        var newMeta = new ColumnMetadata[columns.Length];
        for (int i = 0; i < columns.Length; i++)
        {
            var idx = GetColumnIndex(columns[i]);
            newCols[i] = _columns[idx];
            newMeta[i] = new ColumnMetadata(columns[i], i, _schema.Columns[idx].DataType);
        }
        return new NetDataFrame(columns, newCols, new Schema(newMeta));
    }

    public IDataFrame Filter(Func<DataRow, bool> predicate)
    {
        Guard.NotNull(predicate);
        var matching = new List<int>(RowCount / 2);
        for (int r = 0; r < RowCount; r++)
            if (predicate(BuildRow(r))) matching.Add(r);
        return SliceRows(matching);
    }

    public IDataFrame SortBy(string column, bool ascending = true)
    {
        var idx = GetColumnIndex(column);
        var col = _columns[idx];
        var rowIndices = Enumerable.Range(0, RowCount).ToArray();
        Array.Sort(rowIndices, (a, b) =>
        {
            int cmp = CompareValues(col[a], col[b]);
            return ascending ? cmp : -cmp;
        });
        return SliceRows(rowIndices);
    }

    public IDataFrame AddColumn(string name, object? value)
    {
        Guard.NotNullOrWhiteSpace(name);
        var newArray = new object?[RowCount];
        Array.Fill(newArray, value);

        if (_nameIndex.TryGetValue(name, out var existing))
        {
            var newCols2 = (object?[][])_columns.Clone();
            newCols2[existing] = newArray;
            return new NetDataFrame(_columnNames, newCols2, _schema);
        }

        var appendedNames = new string[_columnNames.Length + 1];
        Array.Copy(_columnNames, appendedNames, _columnNames.Length);
        appendedNames[^1] = name;

        var appendedCols = new object?[_columns.Length + 1][];
        Array.Copy(_columns, appendedCols, _columns.Length);
        appendedCols[^1] = newArray;

        var newMeta = _schema.Columns.Append(new ColumnMetadata(name, _columnNames.Length)).ToArray();
        return new NetDataFrame(appendedNames, appendedCols, new Schema(newMeta));
    }

    public IDataFrame RemoveColumn(string name)
    {
        var idx = GetColumnIndex(name);
        var newNames = _columnNames.Where((_, i) => i != idx).ToArray();
        var newCols = _columns.Where((_, i) => i != idx).ToArray();
        var newMeta = _schema.Columns.Where((_, i) => i != idx)
                             .Select((c, i) => new ColumnMetadata(c.Name, i, c.DataType)).ToArray();
        return new NetDataFrame(newNames, newCols, new Schema(newMeta));
    }

    public IReadOnlyDictionary<object?, IDataFrame> GroupBy(string column)
    {
        var idx = GetColumnIndex(column);
        var col = _columns[idx];
        var groups = new Dictionary<object?, List<int>>(EqualityComparer<object?>.Default);
        for (int r = 0; r < RowCount; r++)
        {
            var key = col[r];
            if (!groups.TryGetValue(key, out var list))
                groups[key] = list = [];
            list.Add(r);
        }
        return groups.ToDictionary(kvp => kvp.Key, kvp => (IDataFrame)SliceRows(kvp.Value));
    }

    public void ToExcel(string path, string sheetName = "Sheet1")
        => ExcelWriterResolver.Default.Write(this, path, sheetName);

    public void ToCsv(string path, char delimiter = ',')
        => CsvWriterResolver.Default.Write(this, path, delimiter);

    // ── Enumerator ────────────────────────────────────────────────────────────

    public IEnumerator<DataRow> GetEnumerator()
    {
        for (int r = 0; r < RowCount; r++)
            yield return BuildRow(r);
    }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    // ── Internal helpers ──────────────────────────────────────────────────────

    private int GetColumnIndex(string name)
    {
        if (_nameIndex.TryGetValue(name, out var idx)) return idx;
        throw new KeyNotFoundException(
            $"Column '{name}' not found. Available: {string.Join(", ", _columnNames)}");
    }

    private DataRow BuildRow(int rowIndex)
    {
        var values = new object?[_columnNames.Length];
        for (int c = 0; c < _columnNames.Length; c++)
            values[c] = _columns[c][rowIndex];
        return new DataRow(values, _nameIndex);
    }

    private NetDataFrame SliceRows(IList<int> rowIndices)
    {
        var newCols = new object?[_columnNames.Length][];
        for (int c = 0; c < _columnNames.Length; c++)
        {
            newCols[c] = new object?[rowIndices.Count];
            var src = _columns[c];
            for (int r = 0; r < rowIndices.Count; r++)
                newCols[c][r] = src[rowIndices[r]];
        }
        return new NetDataFrame(_columnNames, newCols, _schema);
    }

    private static int CompareValues(object? a, object? b)
    {
        if (a is null && b is null) return 0;
        if (a is null) return -1;
        if (b is null) return 1;
        if (a is IComparable ca) return ca.CompareTo(b);
        return string.Compare(a.ToString(), b.ToString(), StringComparison.Ordinal);
    }

    // ── Static factories ──────────────────────────────────────────────────────

    /// <summary>Creates a DataFrame from a column-oriented dictionary.</summary>
    public static NetDataFrame FromColumns(Dictionary<string, object?[]> data)
    {
        Guard.NotNull(data);
        if (data.Count == 0) return Empty();
        var firstLen = data.First().Value.Length;
        if (data.Any(kv => kv.Value.Length != firstLen))
            throw new ArgumentException("All column arrays must have the same length.");
        var names = data.Keys.ToArray();
        var cols = data.Values.ToArray();
        var meta = names.Select((n, i) => new ColumnMetadata(n, i)).ToArray();
        return new NetDataFrame(names, cols, new Schema(meta));
    }

    /// <summary>Creates a DataFrame from a strongly-typed list using reflection.</summary>
    public static NetDataFrame FromList<T>(IEnumerable<T> items) where T : class
    {
        Guard.NotNull(items);
        var list = items as IList<T> ?? items.ToList();
        if (list.Count == 0) return Empty();

        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                             .Where(p => p.CanRead).ToArray();
        var cols = new object?[props.Length][];
        for (int c = 0; c < props.Length; c++)
        {
            cols[c] = new object?[list.Count];
            var prop = props[c];
            for (int r = 0; r < list.Count; r++)
                cols[c][r] = prop.GetValue(list[r]);
        }
        var names = props.Select(p => p.Name).ToArray();
        var meta  = props.Select((p, i) => new ColumnMetadata(p.Name, i, NetTypeToDataType(p.PropertyType))).ToArray();
        return new NetDataFrame(names, cols, new Schema(meta));
    }

    /// <summary>Creates a DataFrame from raw string rows (used by CSV and Excel readers).</summary>
    public static NetDataFrame FromRawRows(string[] headers, IList<string[]> rows,
        bool inferTypes = true, int inferSampleSize = 200)
    {
        Guard.NotNull(headers);
        Guard.NotNull(rows);

        int colCount = headers.Length;
        var cols = new object?[colCount][];
        for (int c = 0; c < colCount; c++)
            cols[c] = new object?[rows.Count];

        DataType[] types = new DataType[colCount];
        if (inferTypes)
        {
            int sampleEnd = Math.Min(inferSampleSize, rows.Count);
            for (int c = 0; c < colCount; c++)
            {
                var sample = rows.Take(sampleEnd).Select(r => c < r.Length ? r[c] : null);
                types[c] = TypeInferenceHelper.Infer(sample);
            }
        }
        else
        {
            Array.Fill(types, DataType.String);
        }

        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (int c = 0; c < colCount; c++)
            {
                var raw = c < row.Length ? row[c] : null;
                cols[c][r] = inferTypes
                    ? TypeInferenceHelper.Convert(raw, types[c])
                    : raw;
            }
        }

        var meta = headers.Select((h, i) => new ColumnMetadata(h, i, types[i])).ToArray();
        return new NetDataFrame(headers, cols, new Schema(meta));
    }

    /// <summary>Returns an empty DataFrame.</summary>
    public static NetDataFrame Empty() => new([], [], new Schema([]));

    private static DataType NetTypeToDataType(Type t)
    {
        var u = Nullable.GetUnderlyingType(t) ?? t;
        return u switch
        {
            _ when u == typeof(string)                          => DataType.String,
            _ when u == typeof(bool)                           => DataType.Boolean,
            _ when u == typeof(int)  || u == typeof(long)  ||
                   u == typeof(short)|| u == typeof(byte)      => DataType.Int64,
            _ when u == typeof(float)|| u == typeof(double)    => DataType.Double,
            _ when u == typeof(decimal)                        => DataType.Decimal,
            _ when u == typeof(DateTime)                       => DataType.DateTime,
            _ when u == typeof(DateOnly)                       => DataType.Date,
            _ => DataType.String
        };
    }

    public override string ToString() =>
        $"NetDataFrame [{RowCount} rows × {ColumnCount} cols] " +
        $"Columns: [{string.Join(", ", _columnNames)}]";
}
