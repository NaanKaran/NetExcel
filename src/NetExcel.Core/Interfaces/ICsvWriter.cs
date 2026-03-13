namespace NetXLCsv.Core.Interfaces;

/// <summary>Writes DataFrames to CSV files.</summary>
public interface ICsvWriter
{
    /// <summary>Writes a DataFrame to a CSV file.</summary>
    void Write(IDataFrame dataFrame, string path, char delimiter = ',', System.Text.Encoding? encoding = null);

    /// <summary>Writes a DataFrame to a stream.</summary>
    void Write(IDataFrame dataFrame, Stream stream, char delimiter = ',', System.Text.Encoding? encoding = null);
}
