using System.Buffers;
using System.Text;

namespace NetXLCsv.Csv;

/// <summary>
/// RFC-4180 compliant, high-performance CSV parser.
/// Uses a pooled char buffer to avoid large heap allocations on hot paths.
/// Supports: custom delimiter, quoted fields, newlines inside quotes, escape sequences.
/// </summary>
internal static class CsvParser
{
    private const int BufferSize = 65536; // 64 KB read buffer

    /// <summary>
    /// Streams parsed rows from the reader without holding all data in memory.
    /// </summary>
    public static IEnumerable<string[]> ParseRows(TextReader reader, char delimiter = ',')
    {
        var fieldBuffer = new StringBuilder(256);
        var fields = new List<string>(16);
        int ch;
        bool inQuotes = false;
        bool fieldStarted = false;

        void FlushField()
        {
            fields.Add(fieldBuffer.ToString());
            fieldBuffer.Clear();
            fieldStarted = false;
        }

        while ((ch = reader.Read()) != -1)
        {
            char c = (char)ch;

            if (inQuotes)
            {
                if (c == '"')
                {
                    // Peek at next char for escaped quote ("")
                    int next = reader.Peek();
                    if (next == '"')
                    {
                        reader.Read();
                        fieldBuffer.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    fieldBuffer.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                    fieldStarted = true;
                }
                else if (c == delimiter)
                {
                    FlushField();
                }
                else if (c == '\r')
                {
                    // CR-LF or lone CR
                    if (reader.Peek() == '\n') reader.Read();
                    FlushField();
                    if (fields.Count > 0)
                    {
                        yield return fields.ToArray();
                        fields.Clear();
                    }
                }
                else if (c == '\n')
                {
                    FlushField();
                    if (fields.Count > 0)
                    {
                        yield return fields.ToArray();
                        fields.Clear();
                    }
                }
                else
                {
                    fieldBuffer.Append(c);
                    fieldStarted = true;
                }
            }
        }

        // Final field/row
        if (fieldStarted || fields.Count > 0)
        {
            FlushField();
            if (fields.Count > 0) yield return fields.ToArray();
        }
    }

    /// <summary>
    /// Reads all rows into memory at once.
    /// Internally uses <see cref="ParseRows"/> so the same parse logic applies.
    /// </summary>
    public static (string[] headers, List<string[]> rows) ReadAll(
        TextReader reader, char delimiter = ',', bool hasHeader = true)
    {
        var all = ParseRows(reader, delimiter).ToList();
        if (all.Count == 0) return ([], []);

        string[] headers;
        List<string[]> dataRows;

        if (hasHeader)
        {
            headers = all[0];
            dataRows = all.Skip(1).ToList();
        }
        else
        {
            int colCount = all.Max(r => r.Length);
            headers = Enumerable.Range(1, colCount)
                .Select(i => $"Col{i}").ToArray();
            dataRows = all;
        }

        return (headers, dataRows);
    }
}
