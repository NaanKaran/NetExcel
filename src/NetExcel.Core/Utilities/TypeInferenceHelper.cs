using System.Globalization;
using NetXLCsv.Core.Models;

namespace NetXLCsv.Core.Utilities;

/// <summary>
/// Infers the <see cref="DataType"/> of column values from raw string samples.
/// Designed for fast, allocation-light inference over a sample window.
/// </summary>
public static class TypeInferenceHelper
{
    private static readonly string[] DateFormats =
    [
        "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy",
        "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss",
        "dd-MMM-yyyy", "MMM dd, yyyy"
    ];

    /// <summary>
    /// Infers the best-fit DataType from a collection of raw string samples.
    /// Empty/null values are skipped.
    /// </summary>
    public static DataType Infer(IEnumerable<string?> samples)
    {
        bool canBeBool = true, canBeInt = true, canBeDouble = true,
             canBeDecimal = true, canBeDate = true, canBeDateTime = true;
        int nonEmpty = 0;

        foreach (var raw in samples)
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            nonEmpty++;

            if (canBeBool && !bool.TryParse(raw, out _))
                canBeBool = false;

            if (canBeInt && !long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                canBeInt = false;

            if (canBeDouble && !double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _))
                canBeDouble = false;

            if (canBeDecimal && !decimal.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _))
                canBeDecimal = false;

            if (canBeDate)
            {
                bool parsedDate = DateTime.TryParseExact(raw, DateFormats,
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                if (!parsedDate) canBeDate = false;
            }

            if (canBeDateTime && !DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                canBeDateTime = false;
        }

        if (nonEmpty == 0) return DataType.String;
        if (canBeBool) return DataType.Boolean;
        if (canBeInt) return DataType.Int64;
        if (canBeDecimal && !canBeInt) return DataType.Decimal;
        if (canBeDouble) return DataType.Double;
        if (canBeDate) return DataType.Date;
        if (canBeDateTime) return DataType.DateTime;
        return DataType.String;
    }

    /// <summary>Converts a raw string to the target type, or returns null on failure.</summary>
    public static object? Convert(string? raw, DataType target)
    {
        if (string.IsNullOrEmpty(raw)) return null;

        return target switch
        {
            DataType.Boolean => bool.TryParse(raw, out var b) ? b : null,
            DataType.Int64 => long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l) ? l : null,
            DataType.Double => double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) ? d : null,
            DataType.Decimal => decimal.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var m) ? m : null,
            DataType.Date => DateTime.TryParseExact(raw, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) ? dt.Date : null,
            DataType.DateTime => DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dts) ? dts : null,
            _ => raw
        };
    }
}
