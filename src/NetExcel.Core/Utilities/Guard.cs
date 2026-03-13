using System.Runtime.CompilerServices;

namespace NetXLCsv.Core.Utilities;

/// <summary>Lightweight argument validation helpers.</summary>
public static class Guard
{
    /// <summary>Throws <see cref="ArgumentNullException"/> if <paramref name="value"/> is null.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static T NotNull<T>(T? value, [CallerArgumentExpression(nameof(value))] string? name = null)
        where T : class
    {
        ArgumentNullException.ThrowIfNull(value, name);
        return value;
    }

    /// <summary>Throws if <paramref name="value"/> is null or whitespace.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string NotNullOrWhiteSpace(string? value, [CallerArgumentExpression(nameof(value))] string? name = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(value, name);
        return value;
    }

    /// <summary>Throws <see cref="ArgumentOutOfRangeException"/> if <paramref name="index"/> is negative or &gt;= <paramref name="max"/>.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void IndexInRange(int index, int max, [CallerArgumentExpression(nameof(index))] string? name = null)
    {
        if ((uint)index >= (uint)max)
            throw new ArgumentOutOfRangeException(name, index, $"Index must be in [0, {max}).");
    }
}
