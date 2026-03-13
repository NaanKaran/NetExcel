namespace NetXLCsv.Formatting;

/// <summary>Describes a conditional formatting rule applied to a cell range.</summary>
public sealed class ConditionalFormat
{
    /// <summary>1-based start row of the range.</summary>
    public int StartRow { get; init; }
    /// <summary>1-based end row (inclusive).</summary>
    public int EndRow { get; init; }
    /// <summary>1-based start column.</summary>
    public int StartColumn { get; init; }
    /// <summary>1-based end column (inclusive).</summary>
    public int EndColumn { get; init; }

    /// <summary>The condition that triggers this format.</summary>
    public required ConditionalRule Rule { get; init; }

    /// <summary>Style to apply when the condition is true.</summary>
    public required CellStyle Style { get; init; }
}

/// <summary>Defines a single conditional formatting rule.</summary>
public sealed class ConditionalRule
{
    public ConditionalOperator Operator { get; init; }
    public object? Value1 { get; init; }
    public object? Value2 { get; init; }

    // ── Factory helpers ────────────────────────────────────────────────────────

    public static ConditionalRule GreaterThan(object value) =>
        new() { Operator = ConditionalOperator.GreaterThan, Value1 = value };

    public static ConditionalRule LessThan(object value) =>
        new() { Operator = ConditionalOperator.LessThan, Value1 = value };

    public static ConditionalRule EqualTo(object value) =>
        new() { Operator = ConditionalOperator.Equal, Value1 = value };

    public static ConditionalRule Between(object low, object high) =>
        new() { Operator = ConditionalOperator.Between, Value1 = low, Value2 = high };

    public static ConditionalRule ContainsText(string text) =>
        new() { Operator = ConditionalOperator.ContainsText, Value1 = text };
}

/// <summary>Comparison operators for conditional formatting.</summary>
public enum ConditionalOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Between,
    NotBetween,
    ContainsText,
    NotContainsText,
    BeginsWith,
    EndsWith
}
