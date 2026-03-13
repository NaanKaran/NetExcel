namespace NetXLCsv.Core.Utilities;

/// <summary>Helpers for generating and converting Excel column names.</summary>
public static class ColumnNameHelper
{
    /// <summary>
    /// Converts a 1-based column number to its Excel letter representation.
    /// E.g. 1 → "A", 26 → "Z", 27 → "AA".
    /// </summary>
    public static string NumberToLetter(int column)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        var result = string.Empty;
        while (column > 0)
        {
            column--;
            result = (char)('A' + column % 26) + result;
            column /= 26;
        }
        return result;
    }

    /// <summary>
    /// Converts an Excel column letter to its 1-based column number.
    /// E.g. "A" → 1, "AA" → 27.
    /// </summary>
    public static int LetterToNumber(string letters)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(letters);
        int result = 0;
        foreach (char c in letters.ToUpperInvariant())
        {
            if (c < 'A' || c > 'Z') throw new FormatException($"Invalid column letter: '{c}'");
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    /// <summary>Generates sequential column names: Col1, Col2, … ColN.</summary>
    public static IEnumerable<string> Generate(int count, string prefix = "Col")
    {
        for (int i = 1; i <= count; i++)
            yield return prefix + i;
    }
}
