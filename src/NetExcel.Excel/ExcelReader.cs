using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NetXLCsv.Core;
using NetXLCsv.Core.Interfaces;
using NetXLCsv.Core.Utilities;

namespace NetXLCsv.Excel;

/// <summary>
/// Reads .xlsx files using DocumentFormat.OpenXml.
/// Uses a SAX-style (OpenXmlReader) approach for large files to minimise DOM allocation.
/// </summary>
public sealed class ExcelReader : IExcelReader
{
    /// <inheritdoc/>
    public IWorkbook OpenWorkbook(string path)
    {
        Guard.NotNullOrWhiteSpace(path);
        if (!File.Exists(path)) throw new FileNotFoundException("Excel file not found.", path);

        // Read all sheets into an ExcelWorkbook
        var wb = ExcelWorkbook.Create();
        using var doc = SpreadsheetDocument.Open(path, isEditable: false);
        LoadSheets(doc, wb);
        return wb;
    }

    /// <inheritdoc/>
    public IWorkbook OpenWorkbook(Stream stream)
    {
        Guard.NotNull(stream);
        var wb = ExcelWorkbook.Create();
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        LoadSheets(doc, wb);
        return wb;
    }

    /// <inheritdoc/>
    public IDataFrame ReadDataFrame(string path, string? sheetName = null, bool hasHeader = true)
    {
        Guard.NotNullOrWhiteSpace(path);
        if (!File.Exists(path)) throw new FileNotFoundException("Excel file not found.", path);

        using var doc = SpreadsheetDocument.Open(path, isEditable: false);
        return ReadSheetAsDataFrame(doc, sheetName, hasHeader);
    }

    /// <inheritdoc/>
    public IDataFrame ReadDataFrame(Stream stream, string? sheetName = null, bool hasHeader = true)
    {
        Guard.NotNull(stream);
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        return ReadSheetAsDataFrame(doc, sheetName, hasHeader);
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private static void LoadSheets(SpreadsheetDocument doc, ExcelWorkbook wb)
    {
        var wbPart = doc.WorkbookPart ?? throw new InvalidDataException("Missing WorkbookPart.");
        var sheets = wbPart.Workbook.Sheets?.Cast<Sheet>() ?? [];

        foreach (var sheet in sheets)
        {
            var sheetName = sheet.Name?.Value ?? "Sheet";
            var ws = (ExcelWorksheet)wb.AddWorksheet(sheetName);

            var part = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            var sharedStrings = GetSharedStrings(wbPart);

            int row = 0;
            using var reader = OpenXmlReader.Create(part);
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row))
                {
                    row = (int)((Row)reader.LoadCurrentElement()!).RowIndex!.Value;
                }
                else if (reader.ElementType == typeof(Cell))
                {
                    var cell = (Cell)reader.LoadCurrentElement()!;
                    var (r, c) = ParseCellRef(cell.CellReference?.Value ?? "A1");
                    var val = GetCellValue(cell, sharedStrings);
                    if (val is not null) ws.WriteCell(r, c, val);
                }
            }
        }
    }

    private static IDataFrame ReadSheetAsDataFrame(SpreadsheetDocument doc, string? sheetName, bool hasHeader)
    {
        var wbPart = doc.WorkbookPart ?? throw new InvalidDataException("Missing WorkbookPart.");
        var sheets = wbPart.Workbook.Sheets?.Cast<Sheet>().ToList() ?? [];

        Sheet? target = sheetName is null
            ? sheets.FirstOrDefault()
            : sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

        if (target is null)
            throw new ArgumentException(sheetName is null ? "No sheets found." : $"Sheet '{sheetName}' not found.");

        var part = (WorksheetPart)wbPart.GetPartById(target.Id!.Value!);
        var sharedStrings = GetSharedStrings(wbPart);

        // Collect rows as string arrays
        var rawRows = new List<List<string>>();
        var currentRow = new List<string>();
        int lastRowIndex = -1;

        foreach (var row in part.Worksheet.Descendants<Row>())
        {
            int rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (lastRowIndex >= 0 && rowIdx > lastRowIndex + 1)
            {
                // Gap rows
                for (int g = lastRowIndex + 1; g < rowIdx; g++)
                    rawRows.Add([]);
            }

            currentRow = [];
            int expectedCol = 1;
            foreach (var cell in row.Descendants<Cell>())
            {
                var (_, c) = ParseCellRef(cell.CellReference?.Value ?? "A1");
                while (expectedCol < c) { currentRow.Add(string.Empty); expectedCol++; }
                currentRow.Add(GetCellValue(cell, sharedStrings)?.ToString() ?? string.Empty);
                expectedCol++;
            }
            rawRows.Add(currentRow);
            lastRowIndex = rowIdx;
        }

        if (rawRows.Count == 0) return NetDataFrame.Empty();

        int colCount = rawRows.Max(r => r.Count);
        string[] headers;
        IList<string[]> dataRows;

        if (hasHeader && rawRows.Count > 0)
        {
            var headerRow = rawRows[0];
            headers = Enumerable.Range(0, colCount)
                .Select(i => i < headerRow.Count && !string.IsNullOrWhiteSpace(headerRow[i])
                    ? headerRow[i]
                    : ColumnNameHelper.NumberToLetter(i + 1))
                .ToArray();
            dataRows = rawRows.Skip(1).Select(r => PadRow(r, colCount)).ToArray();
        }
        else
        {
            headers = ColumnNameHelper.Generate(colCount).ToArray();
            dataRows = rawRows.Select(r => PadRow(r, colCount)).ToArray();
        }

        return NetDataFrame.FromRawRows(headers, dataRows);
    }

    private static string[] PadRow(List<string> row, int count)
    {
        var arr = new string[count];
        for (int i = 0; i < count; i++)
            arr[i] = i < row.Count ? row[i] : string.Empty;
        return arr;
    }

    private static List<string> GetSharedStrings(WorkbookPart wbPart)
    {
        var ssp = wbPart.SharedStringTablePart;
        if (ssp is null) return [];
        return ssp.SharedStringTable.Elements<SharedStringItem>()
            .Select(si => si.Text?.Text ?? si.InnerText)
            .ToList();
    }

    private static object? GetCellValue(Cell cell, List<string> sharedStrings)
    {
        var raw = cell.CellValue?.Text;
        if (raw is null) return null;

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            int idx = int.Parse(raw);
            return idx < sharedStrings.Count ? sharedStrings[idx] : raw;
        }

        if (cell.DataType?.Value == CellValues.Boolean)
            return raw == "1";

        if (cell.DataType?.Value == CellValues.Date)
        {
            if (DateTime.TryParse(raw, out var dt)) return dt;
        }

        // Try numeric (OA date or plain number)
        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
        {
            // Check if style suggests it's a date
            return d;
        }

        return raw;
    }

    private static (int row, int col) ParseCellRef(string cellRef)
    {
        // e.g. "AB123"
        int splitIdx = 0;
        while (splitIdx < cellRef.Length && char.IsLetter(cellRef[splitIdx]))
            splitIdx++;

        var colStr = cellRef[..splitIdx];
        var rowStr = cellRef[splitIdx..];

        int col = ColumnNameHelper.LetterToNumber(colStr);
        int row = int.TryParse(rowStr, out var r) ? r : 1;
        return (row, col);
    }
}
