using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
// Alias the OpenXML CellStyle element to avoid ambiguity with our NetXLCsv.Formatting.CellStyle record.
using OxCellStyle  = DocumentFormat.OpenXml.Spreadsheet.CellStyle;
using OxCellStyles = DocumentFormat.OpenXml.Spreadsheet.CellStyles;

namespace NetXLCsv.Formatting;

/// <summary>
/// Builds and manages the OpenXML Stylesheet part.
/// Responsible for registering fonts, fills, borders, cell formats,
/// and returning their integer index for use in cell CellFormat.
/// </summary>
public sealed class StyleSheet
{
    // Indexes into the various pools
    private readonly List<FontStyle> _fonts = [];
    private readonly List<string?> _fills = [];          // background hex colors
    private readonly List<BorderStyle> _borders = [];
    private readonly List<(int fontIdx, int fillIdx, int borderIdx, string? numFmt)> _cellFormats = [];
    private readonly Dictionary<string, int> _numFmtIds = [];
    private int _customNumFmtBase = 164; // OpenXML built-in formats stop at 163

    public StyleSheet()
    {
        // OpenXML requires at minimum: 1 font, 2 fills, 1 border, 1 cell format
        _fonts.Add(FontStyle.Default);
        _fills.Add(null);          // none (required by spec)
        _fills.Add("#C0C0C0");     // gray125 pattern (required by spec)
        _borders.Add(BorderStyle.None);
        _cellFormats.Add((0, 0, 0, null)); // default format index 0
    }

    /// <summary>Registers a CellStyle and returns its format index for use in cells.</summary>
    public int RegisterStyle(CellStyle style)
    {
        int fontIdx = RegisterFont(style.Font);
        int fillIdx = RegisterFill(style.BackgroundColor);
        int borderIdx = RegisterBorder(style.Border);
        string? numFmt = style.NumberFormat;

        var fmt = (fontIdx, fillIdx, borderIdx, numFmt);
        int existing = _cellFormats.IndexOf(fmt);
        if (existing >= 0) return existing;

        _cellFormats.Add(fmt);
        return _cellFormats.Count - 1;
    }

    private int RegisterFont(FontStyle font)
    {
        int idx = _fonts.IndexOf(font);
        if (idx >= 0) return idx;
        _fonts.Add(font);
        return _fonts.Count - 1;
    }

    private int RegisterFill(string? hex)
    {
        int idx = _fills.IndexOf(hex);
        if (idx >= 0) return idx;
        _fills.Add(hex);
        return _fills.Count - 1;
    }

    private int RegisterBorder(BorderStyle border)
    {
        int idx = _borders.IndexOf(border);
        if (idx >= 0) return idx;
        _borders.Add(border);
        return _borders.Count - 1;
    }

    /// <summary>Builds the OpenXML Stylesheet element.</summary>
    public Stylesheet Build()
    {
        var stylesheet = new Stylesheet();

        // ── Number Formats ────────────────────────────────────────────────────
        var numberingFormats = new NumberingFormats();
        foreach (var (fmt, id) in _numFmtIds)
        {
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = (uint)id,
                FormatCode = fmt
            });
        }
        numberingFormats.Count = (uint)_numFmtIds.Count;
        if (_numFmtIds.Count > 0) stylesheet.Append(numberingFormats);

        // ── Fonts ─────────────────────────────────────────────────────────────
        var fonts = new Fonts { Count = (uint)_fonts.Count };
        foreach (var f in _fonts)
        {
            var font = new Font();
            if (f.IsBold) font.Append(new Bold());
            if (f.IsItalic) font.Append(new Italic());
            if (f.IsUnderline) font.Append(new Underline());
            if (f.IsStrikethrough) font.Append(new Strike());
            font.Append(new FontSize { Val = f.Size });
            font.Append(new FontName { Val = f.Name });
            if (f.Color is not null)
                font.Append(new Color { Rgb = NormalizeHex(f.Color) });
            fonts.Append(font);
        }
        stylesheet.Append(fonts);

        // ── Fills ─────────────────────────────────────────────────────────────
        var fills = new Fills { Count = (uint)_fills.Count };
        foreach (var hex in _fills)
        {
            var fill = new Fill();
            if (hex is null)
            {
                fill.Append(new PatternFill { PatternType = PatternValues.None });
            }
            else if (hex == "#C0C0C0")
            {
                fill.Append(new PatternFill { PatternType = PatternValues.Gray125 });
            }
            else
            {
                var pf = new PatternFill { PatternType = PatternValues.Solid };
                pf.Append(new ForegroundColor { Rgb = NormalizeHex(hex) });
                pf.Append(new BackgroundColor { Indexed = 64 });
                fill.Append(pf);
            }
            fills.Append(fill);
        }
        stylesheet.Append(fills);

        // ── Borders ───────────────────────────────────────────────────────────
        var borders = new Borders { Count = (uint)_borders.Count };
        foreach (var b in _borders)
        {
            var border = new Border();
            border.Append(MakeBorderEdge<LeftBorder>(b.Left, b.Color));
            border.Append(MakeBorderEdge<RightBorder>(b.Right, b.Color));
            border.Append(MakeBorderEdge<TopBorder>(b.Top, b.Color));
            border.Append(MakeBorderEdge<BottomBorder>(b.Bottom, b.Color));
            border.Append(new DiagonalBorder());
            borders.Append(border);
        }
        stylesheet.Append(borders);

        // ── Cell Style Formats (required by spec) ─────────────────────────────
        var cellStyleFormats = new CellStyleFormats { Count = 1 };
        cellStyleFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
        stylesheet.Append(cellStyleFormats);

        // ── Cell Formats ──────────────────────────────────────────────────────
        var cellFormats = new CellFormats { Count = (uint)_cellFormats.Count };
        foreach (var (fIdx, fillIdx, bIdx, numFmt) in _cellFormats)
        {
            uint numFmtId = 0;
            if (numFmt is not null)
            {
                if (!_numFmtIds.TryGetValue(numFmt, out var nid))
                {
                    nid = _customNumFmtBase++;
                    _numFmtIds[numFmt] = nid;
                }
                numFmtId = (uint)nid;
            }
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = numFmtId,
                FontId = (uint)fIdx,
                FillId = (uint)fillIdx,
                BorderId = (uint)bIdx,
                ApplyFont = fIdx > 0,
                ApplyFill = fillIdx > 1,
                ApplyBorder = bIdx > 0,
                ApplyNumberFormat = numFmtId > 0
            });
        }
        stylesheet.Append(cellFormats);

        // ── Cell Styles (required) ────────────────────────────────────────────
        var cellStyles = new OxCellStyles { Count = 1 };
        cellStyles.Append(new OxCellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 });
        stylesheet.Append(cellStyles);

        return stylesheet;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static T MakeBorderEdge<T>(BorderLineStyle style, string? color)
        where T : BorderPropertiesType, new()
    {
        var edge = new T();
        if (style != BorderLineStyle.None)
        {
            edge.Style = style switch
            {
                BorderLineStyle.Thin => BorderStyleValues.Thin,
                BorderLineStyle.Medium => BorderStyleValues.Medium,
                BorderLineStyle.Thick => BorderStyleValues.Thick,
                BorderLineStyle.Dashed => BorderStyleValues.Dashed,
                BorderLineStyle.Dotted => BorderStyleValues.Dotted,
                BorderLineStyle.Double => BorderStyleValues.Double,
                _ => BorderStyleValues.Thin
            };
            if (color is not null)
                edge.Append(new Color { Rgb = NormalizeHex(color) });
        }
        return edge;
    }

    private static string NormalizeHex(string hex) =>
        hex.TrimStart('#').PadLeft(8, 'F').ToUpperInvariant();
}
