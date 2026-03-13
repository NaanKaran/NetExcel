using System.Text;
using NetXLCsv.Core;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Csv;

/// <summary>
/// Reads CSV files into DataFrames using the fast <see cref="CsvParser"/>.
/// </summary>
public sealed class CsvReader : ICsvReader
{
    /// <inheritdoc/>
    public IDataFrame Read(string path, char delimiter = ',', bool hasHeader = true,
        Encoding? encoding = null)
    {
        Guard.NotNullOrWhiteSpace(path);
        if (!File.Exists(path))
            throw new FileNotFoundException("CSV file not found.", path);

        using var stream = File.OpenRead(path);
        return Read(stream, delimiter, hasHeader, encoding);
    }

    /// <inheritdoc/>
    public IDataFrame Read(Stream stream, char delimiter = ',', bool hasHeader = true,
        Encoding? encoding = null)
    {
        Guard.NotNull(stream);
        using var reader = new StreamReader(stream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        var (headers, rows) = CsvParser.ReadAll(reader, delimiter, hasHeader);

        if (headers.Length == 0) return NetDataFrame.Empty();

        return NetDataFrame.FromRawRows(headers, rows, inferTypes: true);
    }

    /// <inheritdoc/>
    public IEnumerable<string[]> StreamRows(string path, char delimiter = ',', bool skipHeader = false)
    {
        Guard.NotNullOrWhiteSpace(path);
        // Uses 'yield return' semantics — caller iterates lazily
        using var stream = File.OpenRead(path);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

        bool first = true;
        foreach (var row in CsvParser.ParseRows(reader, delimiter))
        {
            if (first && skipHeader) { first = false; continue; }
            first = false;
            yield return row;
        }
    }
}
