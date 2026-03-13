namespace NetXLCsv.Streaming;

// Local copy to avoid circular project dependency from Streaming → Core
internal static class ColumnNameHelper
{
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
}
