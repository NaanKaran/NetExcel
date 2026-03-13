using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Models;
using NetXLCsv.Core.Utilities;
using NetXLCsv.Formatting;
// Aliases to resolve CS0104 ambiguities with DocumentFormat.OpenXml.Spreadsheet types
using CellStyle = NetXLCsv.Formatting.CellStyle;
using ConditionalFormat = NetXLCsv.Formatting.ConditionalFormat;

namespace NetXLCsv.Excel;

/// <summary>
/// High-level Excel workbook that manages worksheets and serializes to .xlsx.
/// Uses DocumentFormat.OpenXml (Microsoft's official OpenXML SDK) for robust,
/// standards-compliant output.
/// </summary>
public sealed class ExcelWorkbook : IWorkbook
{
    private readonly List<ExcelWorksheet> _worksheets = [];
    private bool _disposed;

    private ExcelWorkbook() { }

    // ── IWorkbook ─────────────────────────────────────────────────────────────

    /// <inheritdoc/>
    public IReadOnlyList<IWorksheet> Worksheets => _worksheets;

    /// <inheritdoc/>
    public IWorksheet AddWorksheet(string name)
    {
        Guard.NotNullOrWhiteSpace(name);
        if (_worksheets.Any(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException($"A worksheet named '{name}' already exists.");

        var ws = new ExcelWorksheet(name);
        _worksheets.Add(ws);
        return ws;
    }

    /// <inheritdoc/>
    public IWorksheet GetWorksheet(string name)
    {
        return _worksheets.FirstOrDefault(w =>
                   string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase))
               ?? throw new KeyNotFoundException($"Worksheet '{name}' not found.");
    }

    /// <inheritdoc/>
    public void Save(string path)
    {
        Guard.NotNullOrWhiteSpace(path);
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        using var stream = File.Create(path);
        Save(stream);
    }

    /// <inheritdoc/>
    public void Save(Stream stream)
    {
        Guard.NotNull(stream);

        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, autoSave: true);

        // ── WorkbookPart ──────────────────────────────────────────────────────
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // ── Shared strings (improves compression for text-heavy workbooks) ────
        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        var sharedStrings = new SharedStringTable();
        var stringCache = new Dictionary<string, int>(StringComparer.Ordinal);

        // ── Stylesheet ────────────────────────────────────────────────────────
        var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        var styleSheet = new StyleSheet();
        int headerStyleIdx = styleSheet.RegisterStyle(CellStyle.Header);

        // ── Sheets element ────────────────────────────────────────────────────
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        uint sheetId = 1;

        foreach (var ws in _worksheets)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheet = new Worksheet();
            var sheetData = new SheetData();

            // ── Column widths ─────────────────────────────────────────────────
            if (ws.ColumnWidths.Count > 0)
            {
                var cols = new Columns();
                foreach (var (colNum, width) in ws.ColumnWidths.OrderBy(kv => kv.Key))
                {
                    cols.Append(new Column
                    {
                        Min = (uint)colNum,
                        Max = (uint)colNum,
                        Width = width,
                        CustomWidth = true
                    });
                }
                worksheet.Append(cols);
            }

            // ── Cell data ─────────────────────────────────────────────────────
            int maxRow = ws.MaxRow;
            int maxCol = ws.MaxColumn;

            for (int r = 1; r <= maxRow; r++)
            {
                var rowElement = new Row { RowIndex = (uint)r };

                for (int c = 1; c <= maxCol; c++)
                {
                    if (!ws.Cells.TryGetValue((r, c), out var cell)) continue;

                    var cellRef = new CellAddress(r - 1, c - 1).ToA1Notation();
                    var openXmlCell = new Cell { CellReference = cellRef };

                    // Style index
                    int styleIdx = cell.Style is not null
                        ? styleSheet.RegisterStyle(cell.Style)
                        : (r == 1 && cell.Style is null ? headerStyleIdx : 0);
                    if (styleIdx > 0)
                        openXmlCell.StyleIndex = (uint)styleIdx;

                    // Value
                    WriteCellValue(openXmlCell, cell.Value, sharedStrings, stringCache);

                    rowElement.Append(openXmlCell);
                }

                sheetData.Append(rowElement);
            }

            worksheet.Append(sheetData);

            // ── Conditional formats ───────────────────────────────────────────
            foreach (var cf in ws.ConditionalFormats)
                AppendConditionalFormat(worksheet, cf, styleSheet);

            worksheetPart.Worksheet = worksheet;

            // ── Sheet reference ───────────────────────────────────────────────
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = ws.Name
            };
            sheets.Append(sheet);
        }

        // Finalize stylesheet and shared strings
        workbookStylesPart.Stylesheet = styleSheet.Build();
        sharedStringPart.SharedStringTable = sharedStrings;

        workbookPart.Workbook.Save();
    }

    // ── Static factory ────────────────────────────────────────────────────────

    /// <summary>Creates a new empty workbook.</summary>
    public static ExcelWorkbook Create() => new();

    /// <summary>Opens an existing .xlsx file (read mode).</summary>
    public static ExcelWorkbook Open(string path)
    {
        // Reading is handled by ExcelReader; this method is a convenience passthrough
        var reader = new ExcelReader();
        return (ExcelWorkbook)reader.OpenWorkbook(path);
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private static void WriteCellValue(
        Cell cell, object? value,
        SharedStringTable sharedStrings,
        Dictionary<string, int> cache)
    {
        if (value is null) return;

        switch (value)
        {
            case bool b:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(b ? "1" : "0");
                break;

            case DateTime dt:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dt.ToOADate().ToString("R"));
                break;

            case DateOnly d:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(d.ToDateTime(TimeOnly.MinValue).ToOADate().ToString("R"));
                break;

            case float f:
            case double d2:
            case decimal dec:
            case int i:
            case long l:
            case short s:
            case byte by:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDouble(value).ToString("R"));
                break;

            default:
                var text = value.ToString() ?? string.Empty;
                if (!cache.TryGetValue(text, out var idx))
                {
                    idx = cache.Count;
                    cache[text] = idx;
                    sharedStrings.Append(new SharedStringItem(new Text(text)));
                }
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(idx.ToString());
                break;
        }
    }

    private static void AppendConditionalFormat(
        Worksheet ws, ConditionalFormat cf, StyleSheet styleSheet)
    {
        // Build the sqref (range reference)
        var from = new CellAddress(cf.StartRow - 1, cf.StartColumn - 1).ToA1Notation();
        var to = new CellAddress(cf.EndRow - 1, cf.EndColumn - 1).ToA1Notation();
        var sqref = from == to ? from : $"{from}:{to}";

        var cfElement = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue> { InnerText = sqref } };

        var rule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.CellIs,
            FormatId = (uint)styleSheet.RegisterStyle(cf.Style),
            Priority = 1,
            Operator = MapOperator(cf.Rule.Operator)
        };

        if (cf.Rule.Value1 is not null)
            rule.Append(new Formula(cf.Rule.Value1.ToString() ?? string.Empty));
        if (cf.Rule.Value2 is not null)
            rule.Append(new Formula(cf.Rule.Value2.ToString() ?? string.Empty));

        cfElement.Append(rule);
        ws.Append(cfElement);
    }

    private static ConditionalFormattingOperatorValues MapOperator(ConditionalOperator op) => op switch
    {
        ConditionalOperator.Equal => ConditionalFormattingOperatorValues.Equal,
        ConditionalOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
        ConditionalOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
        ConditionalOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
        ConditionalOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
        ConditionalOperator.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
        ConditionalOperator.Between => ConditionalFormattingOperatorValues.Between,
        ConditionalOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
        _ => ConditionalFormattingOperatorValues.Equal
    };

    // ── IDisposable ───────────────────────────────────────────────────────────
    public void Dispose()
    {
        if (!_disposed)
        {
            _worksheets.Clear();
            _disposed = true;
        }
    }
}
