namespace NetXLCsv.Formatting;

/// <summary>Horizontal text alignment within a cell.</summary>
public enum HorizontalAlignment { General, Left, Center, Right, Fill, Justify }

/// <summary>Vertical text alignment within a cell.</summary>
public enum VerticalAlignment { Top, Middle, Bottom }

/// <summary>
/// Encapsulates all formatting attributes for a single cell or range.
/// Immutable record — build new styles using <c>with</c> expressions.
/// </summary>
public sealed record CellStyle
{
    /// <summary>Font settings.</summary>
    public FontStyle Font { get; init; } = FontStyle.Default;

    /// <summary>Fill (background) color in HTML hex format, e.g. "#EEEEEE".</summary>
    public string? BackgroundColor { get; init; }

    /// <summary>Foreground (text) color in HTML hex format.</summary>
    public string? ForegroundColor { get; init; }

    /// <summary>Number format string (e.g. "0.00", "#,##0", "yyyy-MM-dd").</summary>
    public string? NumberFormat { get; init; }

    /// <summary>Horizontal alignment.</summary>
    public HorizontalAlignment HorizontalAlignment { get; init; } = HorizontalAlignment.General;

    /// <summary>Vertical alignment.</summary>
    public VerticalAlignment VerticalAlignment { get; init; } = VerticalAlignment.Bottom;

    /// <summary>Whether text should wrap within the cell.</summary>
    public bool WrapText { get; init; }

    /// <summary>Border settings.</summary>
    public BorderStyle Border { get; init; } = BorderStyle.None;

    // ── Convenience factory ────────────────────────────────────────────────────

    /// <summary>A plain, unstyled cell.</summary>
    public static CellStyle Default { get; } = new();

    /// <summary>Bold header style with light gray background.</summary>
    public static CellStyle Header { get; } = new()
    {
        Font = FontStyle.Bold,
        BackgroundColor = "#EEEEEE",
        HorizontalAlignment = HorizontalAlignment.Center
    };

    /// <summary>Creates a style with the given background color.</summary>
    public static CellStyle WithBackground(string hex) => new() { BackgroundColor = hex };
}

/// <summary>Font formatting settings.</summary>
public sealed record FontStyle
{
    public static FontStyle Default { get; } = new();
    public static FontStyle Bold { get; } = new() { IsBold = true };
    public static FontStyle Italic { get; } = new() { IsItalic = true };
    public static FontStyle BoldItalic { get; } = new() { IsBold = true, IsItalic = true };

    public string Name { get; init; } = "Calibri";
    public double Size { get; init; } = 11;
    public bool IsBold { get; init; }
    public bool IsItalic { get; init; }
    public bool IsUnderline { get; init; }
    public bool IsStrikethrough { get; init; }
    public string? Color { get; init; }
}

/// <summary>Border line style.</summary>
public enum BorderLineStyle { None, Thin, Medium, Thick, Dashed, Dotted, Double }

/// <summary>Border configuration for all four cell edges.</summary>
public sealed record BorderStyle
{
    public static BorderStyle None { get; } = new();
    public static BorderStyle All { get; } = new()
    {
        Top = BorderLineStyle.Thin, Bottom = BorderLineStyle.Thin,
        Left = BorderLineStyle.Thin, Right = BorderLineStyle.Thin
    };
    public static BorderStyle Box { get; } = new()
    {
        Top = BorderLineStyle.Medium, Bottom = BorderLineStyle.Medium,
        Left = BorderLineStyle.Medium, Right = BorderLineStyle.Medium
    };

    public BorderLineStyle Top { get; init; } = BorderLineStyle.None;
    public BorderLineStyle Bottom { get; init; } = BorderLineStyle.None;
    public BorderLineStyle Left { get; init; } = BorderLineStyle.None;
    public BorderLineStyle Right { get; init; } = BorderLineStyle.None;
    public string? Color { get; init; }
}
