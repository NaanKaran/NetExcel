using NetXLCsv.Core.Interfaces;

namespace NetXLCsv.Excel;

/// <summary>
/// Default <see cref="IExcelWriter"/> implementation.
/// Creates a workbook with a single worksheet and writes the DataFrame to it.
/// </summary>
public sealed class ExcelWriter : IExcelWriter
{
    /// <inheritdoc/>
    public void Write(IDataFrame dataFrame, string path, string sheetName = "Sheet1")
    {
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet(sheetName);
        ws.WriteTable(1, 1, dataFrame);
        wb.Save(path);
    }

    /// <inheritdoc/>
    public void Write(IDataFrame dataFrame, Stream stream, string sheetName = "Sheet1")
    {
        using var wb = ExcelWorkbook.Create();
        var ws = (ExcelWorksheet)wb.AddWorksheet(sheetName);
        ws.WriteTable(1, 1, dataFrame);
        wb.Save(stream);
    }
}
