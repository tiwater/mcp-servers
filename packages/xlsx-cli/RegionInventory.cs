using System.Security.Cryptography;

namespace Dockit.Xlsx;

public sealed record XlsxRegionInventory(
    string Schema,
    string ToolVersion,
    string File,
    string InputSha256,
    IReadOnlyList<XlsxRegionInventorySheet> Sheets);

public sealed record XlsxRegionInventorySheet(
    string Name,
    IReadOnlyList<XlsxRegionInventoryRegion> Regions);

public sealed record XlsxRegionInventoryRegion(
    string Id,
    string Range,
    int StartRow,
    int EndRow,
    IReadOnlyList<XlsxRegionInventoryRow> Rows);

public sealed record XlsxRegionInventoryRow(
    int Row,
    IReadOnlyList<XlsxRegionInventoryCell> Cells);

public sealed record XlsxRegionInventoryCell(
    string Reference,
    int Row,
    int Column,
    string ColumnName,
    string RawValue,
    string FormattedValue,
    string? Formula);

public sealed record XlsxRegionInventoryV2(
    string Schema,
    string ToolVersion,
    string File,
    string InputSha256,
    IReadOnlyList<XlsxRegionInventorySheetV2> Sheets);

public sealed record XlsxRegionInventorySheetV2(
    string Name,
    IReadOnlyList<XlsxRegionInventoryRegionV2> Regions);

public sealed record XlsxRegionInventoryRegionV2(
    string Id,
    string Range,
    int StartRow,
    int EndRow,
    IReadOnlyList<XlsxRegionInventoryRowV2> Rows);

public sealed record XlsxRegionInventoryRowV2(
    int Row,
    IReadOnlyList<XlsxRegionInventoryCellV2> Cells);

public sealed record XlsxRegionInventoryCellV2(
    string Reference,
    int Row,
    int Column,
    string ColumnName,
    string RawValue,
    string FormattedValue,
    string? Formula,
    XlsxNormalizedValue NormalizedValue);

public static class RegionInventory
{
    public static XlsxRegionInventory Inspect(string path)
    {
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
            throw new FileNotFoundException("Workbook not found.", fullPath);

        var workbook = WorkbookLoader.Load(fullPath);
        var sheets = workbook.Sheets.Select((sheet, sheetIndex) =>
        {
            var materialCells = sheet.Cells
                .Where(IsMaterial)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();
            var regions = BuildRegions(materialCells, sheetIndex + 1);
            return new XlsxRegionInventorySheet(sheet.Name, regions);
        }).ToList();

        return new XlsxRegionInventory(
            "tiwater.xlsx.region-inventory/v1",
            XlsxToolVersion.Current,
            fullPath,
            System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(fullPath))).ToLowerInvariant(),
            sheets);
    }

    public static XlsxRegionInventoryV2 InspectV2(string path)
    {
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
            throw new FileNotFoundException("Workbook not found.", fullPath);

        var workbook = WorkbookLoader.Load(fullPath, includeNormalizedValues: true);
        var sheets = workbook.Sheets.Select((sheet, sheetIndex) =>
        {
            var materialCells = sheet.Cells
                .Where(IsMaterial)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();
            var regions = BuildRegionsV2(materialCells, sheetIndex + 1);
            return new XlsxRegionInventorySheetV2(sheet.Name, regions);
        }).ToList();

        return new XlsxRegionInventoryV2(
            "tiwater.xlsx.region-inventory/v2",
            XlsxToolVersion.Current,
            fullPath,
            System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(fullPath))).ToLowerInvariant(),
            sheets);
    }

    public static int Run(string[] args)
    {
        if (args.Length < 1)
            throw new InvalidOperationException("inventory-regions requires <input.xlsx> [<output.json>] [--schema v1|v2]");

        var input = args[0];
        string? output = null;
        var schema = "v1";
        for (var index = 1; index < args.Length; index++)
        {
            if (args[index] == "--schema")
            {
                if (++index >= args.Length || args[index] is not ("v1" or "v2"))
                    throw new InvalidOperationException("inventory-regions --schema requires v1 or v2");
                schema = args[index];
                continue;
            }
            if (args[index].StartsWith("--", StringComparison.Ordinal) || output is not null)
                throw new InvalidOperationException($"inventory-regions argument is invalid: {args[index]}");
            output = args[index];
        }

        var json = schema == "v2"
            ? System.Text.Json.JsonSerializer.Serialize(InspectV2(input), Json.Options)
            : System.Text.Json.JsonSerializer.Serialize(Inspect(input), Json.Options);
        if (output is not null)
        {
            var fullOutput = Path.GetFullPath(output);
            File.WriteAllText(fullOutput, json);
            Console.WriteLine(fullOutput);
        }
        else
        {
            Console.WriteLine(json);
        }
        return 0;
    }

    private static bool IsMaterial(WorkbookLoader.CellDataModel cell) =>
        !string.IsNullOrWhiteSpace(cell.Value)
        || !string.IsNullOrWhiteSpace(cell.FormattedValue)
        || !string.IsNullOrWhiteSpace(cell.Formula);

    private static IReadOnlyList<XlsxRegionInventoryRegion> BuildRegions(
        IReadOnlyList<WorkbookLoader.CellDataModel> cells,
        int sheetOrdinal)
    {
        var rows = cells.GroupBy(cell => cell.Row).OrderBy(group => group.Key).ToList();
        if (rows.Count == 0) return [];

        var bands = new List<List<IGrouping<int, WorkbookLoader.CellDataModel>>>();
        foreach (var row in rows)
        {
            if (bands.Count == 0 || row.Key > bands[^1][^1].Key + 1)
                bands.Add([]);
            bands[^1].Add(row);
        }

        return bands.Select((band, regionIndex) =>
        {
            var bandCells = band.SelectMany(group => group).ToList();
            var minColumn = bandCells.Min(cell => cell.Column);
            var maxColumn = bandCells.Max(cell => cell.Column);
            var startRow = band[0].Key;
            var endRow = band[^1].Key;
            var inventoryRows = band.Select(group => new XlsxRegionInventoryRow(
                group.Key,
                group.OrderBy(cell => cell.Column).Select(ToPublishedCell).ToList())).ToList();
            return new XlsxRegionInventoryRegion(
                $"s{sheetOrdinal}-r{regionIndex + 1}",
                $"{ColumnName(minColumn)}{startRow}:{ColumnName(maxColumn)}{endRow}",
                startRow,
                endRow,
                inventoryRows);
        }).ToList();
    }

    private static IReadOnlyList<XlsxRegionInventoryRegionV2> BuildRegionsV2(
        IReadOnlyList<WorkbookLoader.CellDataModel> cells,
        int sheetOrdinal)
    {
        var rows = cells.GroupBy(cell => cell.Row).OrderBy(group => group.Key).ToList();
        if (rows.Count == 0) return [];

        var bands = new List<List<IGrouping<int, WorkbookLoader.CellDataModel>>>();
        foreach (var row in rows)
        {
            if (bands.Count == 0 || row.Key > bands[^1][^1].Key + 1)
                bands.Add([]);
            bands[^1].Add(row);
        }

        return bands.Select((band, regionIndex) =>
        {
            var bandCells = band.SelectMany(group => group).ToList();
            var minColumn = bandCells.Min(cell => cell.Column);
            var maxColumn = bandCells.Max(cell => cell.Column);
            var startRow = band[0].Key;
            var endRow = band[^1].Key;
            var inventoryRows = band.Select(group => new XlsxRegionInventoryRowV2(
                group.Key,
                group.OrderBy(cell => cell.Column).Select(ToPublishedCellV2).ToList())).ToList();
            return new XlsxRegionInventoryRegionV2(
                $"s{sheetOrdinal}-r{regionIndex + 1}",
                $"{ColumnName(minColumn)}{startRow}:{ColumnName(maxColumn)}{endRow}",
                startRow,
                endRow,
                inventoryRows);
        }).ToList();
    }

    private static XlsxRegionInventoryCell ToPublishedCell(WorkbookLoader.CellDataModel cell) => new(
        cell.Reference,
        cell.Row,
        cell.Column,
        ColumnName(cell.Column),
        cell.Value,
        cell.FormattedValue,
        cell.Formula);

    private static XlsxRegionInventoryCellV2 ToPublishedCellV2(WorkbookLoader.CellDataModel cell) => new(
        cell.Reference,
        cell.Row,
        cell.Column,
        ColumnName(cell.Column),
        cell.Value,
        cell.FormattedValue,
        cell.Formula,
        cell.NormalizedValue ?? throw new InvalidDataException($"Cell normalized value is missing: {cell.Reference}"));

    private static string ColumnName(int column)
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
