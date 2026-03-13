namespace NetXLCsv.Core.Models;

/// <summary>
/// Represents the inferred or declared type of a DataFrame column.
/// </summary>
public enum DataType
{
    /// <summary>Type is not known or could not be inferred.</summary>
    Unknown = 0,
    /// <summary>String / text data.</summary>
    String,
    /// <summary>64-bit integer.</summary>
    Int64,
    /// <summary>64-bit floating point.</summary>
    Double,
    /// <summary>High-precision decimal (e.g. currency).</summary>
    Decimal,
    /// <summary>Boolean (true/false).</summary>
    Boolean,
    /// <summary>Date only (no time component).</summary>
    Date,
    /// <summary>Date and time.</summary>
    DateTime,
    /// <summary>Time span.</summary>
    TimeSpan
}
