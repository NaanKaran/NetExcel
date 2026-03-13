namespace NetXLCsv.Core.Interfaces;

/// <summary>Reads CSV files into DataFrames.</summary>
public interface ICsvReader
{
    /// <summary>Reads a CSV file into a DataFrame.</summary>
    IDataFrame Read(string path, char delimiter = ',', bool hasHeader = true, System.Text.Encoding? encoding = null);

    /// <summary>Reads a CSV stream into a DataFrame.</summary>
    IDataFrame Read(Stream stream, char delimiter = ',', bool hasHeader = true, System.Text.Encoding? encoding = null);

    /// <summary>Streams rows from a large CSV without loading all rows into memory at once.</summary>
    IEnumerable<string[]> StreamRows(string path, char delimiter = ',', bool skipHeader = false);
}
