namespace NetXLCsv.Core.Interfaces;

/// <summary>Writes data to Excel workbooks.</summary>
public interface IExcelWriter
{
    /// <summary>Writes a DataFrame to an Excel file.</summary>
    void Write(IDataFrame dataFrame, string path, string sheetName = "Sheet1");

    /// <summary>Writes a DataFrame to a stream.</summary>
    void Write(IDataFrame dataFrame, Stream stream, string sheetName = "Sheet1");
}
