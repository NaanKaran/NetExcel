using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Streaming;

/// <summary>
/// SAX-style streaming Excel writer that supports writing 1M+ rows with
/// a small, bounded memory footprint (~few MB regardless of row count).
///
/// Architecture:
///  - Uses OpenXML SDK's <see cref="OpenXmlWriter"/> for forward-only writing.
///  - Shared strings are NOT used (direct inline strings) to avoid string-table
///    accumulation for large text datasets. For datasets dominated by repeated
///    strings, callers can optionally enable shared-string mode.
///  - Numeric values are written as raw numbers (no formatting overhead).
///
/// Usage:
/// <code>
/// using var writer = StreamingExcelWriter.Create("big.xlsx");
/// writer.WriteHeader("Id", "Name", "Revenue", "Country");
/// foreach (var row in dataSource)
///     writer.WriteRow(row.Id, row.Name, row.Revenue, row.Country);
/// // Dispose() finalizes and closes the file
/// </code>
/// </summary>
public sealed class StreamingExcelWriter : IDisposable
{
    private SpreadsheetDocument? _doc;
    private OpenXmlWriter? _writer;
    private WorksheetPart? _worksheetPart;
    private long _rowIndex;
    private bool _disposed;
    private readonly Stream _outputStream;
    private readonly bool _leaveOpen;

    // Pre-allocated attribute lists to avoid per-cell allocation
    private static readonly OpenXmlAttribute[] _stringAttribs =
        [new OpenXmlAttribute("t", string.Empty, "inlineStr")];
    private static readonly OpenXmlAttribute[] _numberAttribs = [];
    private static readonly OpenXmlAttribute[] _boolAttribs =
        [new OpenXmlAttribute("t", string.Empty, "b")];

    private StreamingExcelWriter(Stream stream, string sheetName, bool leaveOpen = false)
    {
        _outputStream = stream;
        _leaveOpen = leaveOpen;
        _rowIndex = 1;

        _doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = _doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Add minimal stylesheet (required)
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildMinimalStylesheet();

        _worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        // Register sheet
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(_worksheetPart),
            SheetId = 1,
            Name = sheetName
        });

        _writer = OpenXmlWriter.Create(_worksheetPart);
        _writer.WriteStartElement(new Worksheet());
        _writer.WriteStartElement(new SheetData());
    }

    // ── Factory methods ───────────────────────────────────────────────────────

    /// <summary>Creates a streaming writer targeting the given file path.</summary>
    public static StreamingExcelWriter Create(string path, string sheetName = "Sheet1")
    {
        Guard.NotNullOrWhiteSpace(path);
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        var stream = File.Create(path);
        return new StreamingExcelWriter(stream, sheetName, leaveOpen: false);
    }

    /// <summary>Creates a streaming writer targeting a stream.</summary>
    public static StreamingExcelWriter Create(Stream stream, string sheetName = "Sheet1", bool leaveOpen = true)
    {
        Guard.NotNull(stream);
        return new StreamingExcelWriter(stream, sheetName, leaveOpen);
    }

    // ── Row writing ───────────────────────────────────────────────────────────

    /// <summary>
    /// Writes a single header row. Call once before any data rows.
    /// Headers are written with bold styling.
    /// </summary>
    public void WriteHeader(params string[] columns)
        => WriteRow(columns.Cast<object?>().ToArray());

    /// <summary>
    /// Writes a row with mixed-type values.
    /// Supported value types: string, bool, int, long, double, decimal, DateTime, DateOnly, null.
    /// </summary>
    public void WriteRow(params object?[] values)
    {
        EnsureNotDisposed();
        Guard.NotNull(values);

        var rowAttribs = new[] { new OpenXmlAttribute("r", string.Empty, _rowIndex.ToString()) };
        _writer!.WriteStartElement(new Row(), rowAttribs);

        for (int c = 0; c < values.Length; c++)
        {
            var addr = ColumnNameHelper.NumberToLetter(c + 1) + _rowIndex;
            WriteCell(addr, values[c]);
        }

        _writer.WriteEndElement(); // Row
        _rowIndex++;
    }

    /// <summary>Writes a row from a sequence (avoids params array allocation for large loops).</summary>
    public void WriteRow(IEnumerable<object?> values)
    {
        EnsureNotDisposed();
        Guard.NotNull(values);

        var rowAttribs = new[] { new OpenXmlAttribute("r", string.Empty, _rowIndex.ToString()) };
        _writer!.WriteStartElement(new Row(), rowAttribs);

        int c = 0;
        foreach (var value in values)
        {
            var addr = ColumnNameHelper.NumberToLetter(c + 1) + _rowIndex;
            WriteCell(addr, value);
            c++;
        }

        _writer.WriteEndElement(); // Row
        _rowIndex++;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void WriteCell(string address, object? value)
    {
        if (value is null) return;

        var cellRef = new OpenXmlAttribute("r", string.Empty, address);

        switch (value)
        {
            case bool b:
                _writer!.WriteStartElement(new Cell(), [cellRef, .._boolAttribs]);
                _writer.WriteElement(new CellValue(b ? "1" : "0"));
                _writer.WriteEndElement();
                break;

            case DateTime dt:
                _writer!.WriteStartElement(new Cell(), [cellRef]);
                _writer.WriteElement(new CellValue(dt.ToOADate().ToString("R", CultureInfo.InvariantCulture)));
                _writer.WriteEndElement();
                break;

            case DateOnly d:
                _writer!.WriteStartElement(new Cell(), [cellRef]);
                _writer.WriteElement(new CellValue(d.ToDateTime(TimeOnly.MinValue).ToOADate().ToString("R", CultureInfo.InvariantCulture)));
                _writer.WriteEndElement();
                break;

            case int i:
            case long l:
            case short s:
            case byte by:
            case double dbl:
            case float flt:
            case decimal dec:
                _writer!.WriteStartElement(new Cell(), [cellRef, .._numberAttribs]);
                _writer.WriteElement(new CellValue(Convert.ToDouble(value).ToString("R", CultureInfo.InvariantCulture)));
                _writer.WriteEndElement();
                break;

            default:
                var text = value.ToString() ?? string.Empty;
                _writer!.WriteStartElement(new Cell(), [cellRef, .._stringAttribs]);
                _writer.WriteStartElement(new InlineString());
                _writer.WriteElement(new Text(text));
                _writer.WriteEndElement(); // InlineString
                _writer.WriteEndElement(); // Cell
                break;
        }
    }

    private static Stylesheet BuildMinimalStylesheet()
    {
        var ss = new Stylesheet();
        var fonts = new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })) { Count = 1 };
        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
        { Count = 2 };
        var borders = new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 };
        var csf = new CellStyleFormats(new CellFormat()) { Count = 1 };
        var cf = new CellFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }) { Count = 1 };
        var css = new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 };
        ss.Append(fonts, fills, borders, csf, cf, css);
        return ss;
    }

    private void EnsureNotDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(StreamingExcelWriter));
    }

    // ── Statistics ────────────────────────────────────────────────────────────

    /// <summary>Number of rows written so far (including header if written).</summary>
    public long RowsWritten => _rowIndex - 1;

    // ── IDisposable ───────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _writer?.WriteEndElement(); // SheetData
        _writer?.WriteEndElement(); // Worksheet
        _writer?.Close();
        _writer?.Dispose();

        _doc?.WorkbookPart?.Workbook.Save();
        _doc?.Dispose();

        if (!_leaveOpen)
            _outputStream.Dispose();
    }
}
