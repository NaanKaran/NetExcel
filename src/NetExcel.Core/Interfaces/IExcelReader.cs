namespace NetXLCsv.Core.Interfaces;

/// <summary>Reads Excel workbooks and worksheets.</summary>
public interface IExcelReader
{
    /// <summary>Opens a workbook from the given file path.</summary>
    IWorkbook OpenWorkbook(string path);

    /// <summary>Opens a workbook from a stream.</summary>
    IWorkbook OpenWorkbook(Stream stream);

    /// <summary>
    /// Reads the first worksheet of the Excel file directly into a DataFrame
    /// without loading the entire workbook object model.
    /// </summary>
    IDataFrame ReadDataFrame(string path, string? sheetName = null, bool hasHeader = true);

    /// <summary>
    /// Reads the first worksheet from a stream directly into a DataFrame.
    /// The stream must contain a valid .xlsx file.
    /// </summary>
    IDataFrame ReadDataFrame(Stream stream, string? sheetName = null, bool hasHeader = true);
}
