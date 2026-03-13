using System.Globalization;
using System.Text;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Streaming;

/// <summary>
/// Streaming CSV writer for 1M+ row scenarios.
/// Writes rows directly to a buffered <see cref="StreamWriter"/> without buffering
/// the entire dataset in memory.
///
/// Usage:
/// <code>
/// using var csv = StreamingCsvWriter.Create("big.csv");
/// csv.WriteHeader("Id", "Name", "Country");
/// foreach (var r in rows)
///     csv.WriteRow(r.Id, r.Name, r.Country);
/// </code>
/// </summary>
public sealed class StreamingCsvWriter : IDisposable
{
    private readonly StreamWriter _writer;
    private readonly char _delimiter;
    private bool _disposed;
    private long _rowsWritten;

    private StreamingCsvWriter(StreamWriter writer, char delimiter)
    {
        _writer = writer;
        _delimiter = delimiter;
    }

    // ── Factory ───────────────────────────────────────────────────────────────

    public static StreamingCsvWriter Create(string path, char delimiter = ',',
        Encoding? encoding = null)
    {
        Guard.NotNullOrWhiteSpace(path);
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        var stream = File.Create(path);
        var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(true), 65536, leaveOpen: false);
        return new StreamingCsvWriter(writer, delimiter);
    }

    public static StreamingCsvWriter Create(Stream stream, char delimiter = ',',
        Encoding? encoding = null, bool leaveOpen = true)
    {
        Guard.NotNull(stream);
        var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(true), 65536, leaveOpen);
        return new StreamingCsvWriter(writer, delimiter);
    }

    // ── Row writing ───────────────────────────────────────────────────────────

    public void WriteHeader(params string[] columns)
    {
        EnsureNotDisposed();
        WriteRowInternal(columns.Select(c => (object?)c));
    }

    public void WriteRow(params object?[] values)
    {
        EnsureNotDisposed();
        WriteRowInternal(values);
    }

    public void WriteRow(IEnumerable<object?> values)
    {
        EnsureNotDisposed();
        WriteRowInternal(values);
    }

    private void WriteRowInternal(IEnumerable<object?> values)
    {
        bool first = true;
        foreach (var v in values)
        {
            if (!first) _writer.Write(_delimiter);
            _writer.Write(QuoteField(FormatValue(v)));
            first = false;
        }
        _writer.WriteLine();
        _rowsWritten++;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string FormatValue(object? value) => value switch
    {
        null => string.Empty,
        bool b => b ? "true" : "false",
        DateTime dt => dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        DateOnly d => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        double dbl => dbl.ToString("R", CultureInfo.InvariantCulture),
        float flt => flt.ToString("R", CultureInfo.InvariantCulture),
        decimal dec => dec.ToString(CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };

    private string QuoteField(string field)
    {
        if (field.Contains(_delimiter) || field.Contains('"') ||
            field.Contains('\n') || field.Contains('\r'))
        {
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        }
        return field;
    }

    public long RowsWritten => _rowsWritten;

    private void EnsureNotDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(StreamingCsvWriter));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _writer.Flush();
        _writer.Dispose();
    }
}
