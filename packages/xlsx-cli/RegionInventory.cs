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

    public static int Run(string[] args)
    {
        if (args.Length < 1)
            throw new InvalidOperationException("inventory-regions requires <input.xlsx> [<output.json>]");

        var inventory = Inspect(args[0]);
        var json = System.Text.Json.JsonSerializer.Serialize(inventory, Json.Options);
        if (args.Length > 1)
        {
            var output = Path.GetFullPath(args[1]);
            File.WriteAllText(output, json);
            Console.WriteLine(output);
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

    private static XlsxRegionInventoryCell ToPublishedCell(WorkbookLoader.CellDataModel cell) => new(
        cell.Reference,
        cell.Row,
        cell.Column,
        ColumnName(cell.Column),
        cell.Value,
        cell.FormattedValue,
        cell.Formula);

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
