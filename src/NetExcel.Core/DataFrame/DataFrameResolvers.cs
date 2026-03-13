using NetXLCsv.Core.Interfaces;

namespace NetXLCsv.Core;

/// <summary>
/// Lazy resolver for the default IExcelWriter.
/// Uses late binding (Type.GetType) so NetXLCsv.Core has NO compile-time dependency
/// on NetXLCsv.Excel — avoiding circular project references.
/// </summary>
internal static class ExcelWriterResolver
{
    private static IExcelWriter? _default;

    public static IExcelWriter Default
    {
        get
        {
            if (_default is not null) return _default;
            var type = Type.GetType("NetXLCsv.Excel.ExcelWriter, NetXLCsv.Excel")
                ?? throw new InvalidOperationException(
                    "NetXLCsv.Excel assembly not found. Add a reference to NetXLCsv.Excel or " +
                    "call ExcelWriterResolver.Register(yourWriter) from your DI root.");
            _default = (IExcelWriter)Activator.CreateInstance(type)!;
            return _default;
        }
    }

    public static void Register(IExcelWriter writer) => _default = writer;
}

/// <summary>Lazy resolver for the default ICsvWriter.</summary>
internal static class CsvWriterResolver
{
    private static ICsvWriter? _default;

    public static ICsvWriter Default
    {
        get
        {
            if (_default is not null) return _default;
            var type = Type.GetType("NetXLCsv.Csv.CsvWriter, NetXLCsv.Csv")
                ?? throw new InvalidOperationException(
                    "NetXLCsv.Csv assembly not found. Add a reference to NetXLCsv.Csv or " +
                    "call CsvWriterResolver.Register(yourWriter) from your DI root.");
            _default = (ICsvWriter)Activator.CreateInstance(type)!;
            return _default;
        }
    }

    public static void Register(ICsvWriter writer) => _default = writer;
}
