using System.Globalization;
using System.Text;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Csv;

/// <summary>
/// Writes DataFrames to CSV files.
/// Uses a <see cref="StringBuilder"/> field buffer for zero-allocation quoting of simple strings.
/// </summary>
public sealed class CsvWriter : ICsvWriter
{
    /// <inheritdoc/>
    public void Write(IDataFrame dataFrame, string path, char delimiter = ',',
        Encoding? encoding = null)
    {
        Guard.NotNull(dataFrame);
        Guard.NotNullOrWhiteSpace(path);

        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        using var stream = File.Create(path);
        Write(dataFrame, stream, delimiter, encoding);
    }

    /// <inheritdoc/>
    public void Write(IDataFrame dataFrame, Stream stream, char delimiter = ',',
        Encoding? encoding = null)
    {
        Guard.NotNull(dataFrame);
        Guard.NotNull(stream);

        using var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: true), bufferSize: 65536, leaveOpen: true);

        // Header
        WriteRow(writer, dataFrame.Schema.Columns.Select(c => c.Name), delimiter);

        // Data
        foreach (var row in dataFrame)
        {
            var values = Enumerable.Range(0, dataFrame.ColumnCount)
                .Select(i => FormatValue(row[i]));
            WriteRow(writer, values, delimiter);
        }

        writer.Flush();
    }

    private static void WriteRow(TextWriter writer, IEnumerable<string> fields, char delimiter)
    {
        bool first = true;
        foreach (var field in fields)
        {
            if (!first) writer.Write(delimiter);
            writer.Write(QuoteField(field, delimiter));
            first = false;
        }
        writer.WriteLine();
    }

    private static string FormatValue(object? value) => value switch
    {
        null => string.Empty,
        bool b => b ? "true" : "false",
        DateTime dt => dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        DateOnly d => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        float f => f.ToString("R", CultureInfo.InvariantCulture),
        decimal m => m.ToString(CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };

    private static string QuoteField(string field, char delimiter)
    {
        bool needsQuote = field.Contains(delimiter) ||
                          field.Contains('"') ||
                          field.Contains('\n') ||
                          field.Contains('\r');
        if (!needsQuote) return field;

        return "\"" + field.Replace("\"", "\"\"") + "\"";
    }
}
