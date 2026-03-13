namespace NetXLCsv.Core.Interfaces;

/// <summary>Represents an Excel workbook.</summary>
public interface IWorkbook : IDisposable
{
    /// <summary>All worksheets in the workbook.</summary>
    IReadOnlyList<IWorksheet> Worksheets { get; }

    /// <summary>Adds a new worksheet with the given name.</summary>
    IWorksheet AddWorksheet(string name);

    /// <summary>Returns the worksheet with the given name (case-insensitive).</summary>
    IWorksheet GetWorksheet(string name);

    /// <summary>Saves the workbook to the specified file path.</summary>
    void Save(string path);

    /// <summary>Saves the workbook to the provided stream.</summary>
    void Save(Stream stream);
}
