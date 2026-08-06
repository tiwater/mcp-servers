using System.Security.Cryptography;
using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class RegionInventoryTests
{
    [Fact]
    public void Inventory_preserves_distinct_regions_coordinates_formatted_values_and_formulas()
    {
        var path = CreateWorkbook(
            R(2, ("C", "Instrument", null), ("D", "Reader-Z", null)),
            R(6, ("B", "Specimen", null), ("D", "Reported value", null), ("F", "Unit", null)),
            R(7, ("B", "BX-17", null), ("D", "0.024", null), ("F", "EU/mL", null)),
            R(8, ("B", "BX-18", null), ("D", "0.031", null), ("F", "EU/mL", null)),
            R(12, ("A", "Curve", null), ("B", "Response", null)),
            R(13, ("A", "1", null), ("B", "2", "A13*2")));

        var inventory = RegionInventory.Inspect(path);

        Assert.Equal("tiwater.xlsx.region-inventory/v1", inventory.Schema);
        Assert.Equal(SHA256.HashData(File.ReadAllBytes(path)).Select(b => b.ToString("x2")).Aggregate(string.Concat), inventory.InputSha256);
        var sheet = Assert.Single(inventory.Sheets);
        Assert.Collection(sheet.Regions,
            region =>
            {
                Assert.Equal("C2:D2", region.Range);
                Assert.Equal([2], region.Rows.Select(row => row.Row));
            },
            region =>
            {
                Assert.Equal("B6:F8", region.Range);
                Assert.Equal([6, 7, 8], region.Rows.Select(row => row.Row));
                Assert.Contains(region.Rows[1].Cells, cell =>
                    cell.Reference == "D7" && cell.RawValue == "0.024" && cell.FormattedValue == "0.024");
            },
            region =>
            {
                Assert.Equal("A12:B13", region.Range);
                Assert.Contains(region.Rows[1].Cells, cell =>
                    cell.Reference == "B13" && cell.Formula == "A13*2" && cell.RawValue == "2");
            });
    }

    [Fact]
    public void Inventory_is_layout_agnostic_and_keeps_empty_rows_inside_no_region()
    {
        var path = CreateWorkbook(
            R(4, ("H", "Unit", null), ("B", "Result", null), ("E", "Sample", null)),
            R(5, ("H", "mg/L", null), ("B", "<0.5", null), ("E", "Q-9", null)),
            R(9, ("K", "Only", null)));

        var inventory = RegionInventory.Inspect(path);

        var regions = Assert.Single(inventory.Sheets).Regions;
        Assert.Equal(2, regions.Count);
        Assert.Equal("B4:H5", regions[0].Range);
        Assert.Equal(["B4", "E4", "H4"], regions[0].Rows[0].Cells.Select(cell => cell.Reference));
        Assert.Equal("K9:K9", regions[1].Range);
    }

    [Fact]
    public void Inventory_empty_workbook_has_no_regions()
    {
        var inventory = RegionInventory.Inspect(CreateWorkbook());

        Assert.Empty(Assert.Single(inventory.Sheets).Regions);
    }

    [Fact]
    public void Inventory_fails_for_missing_input()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.xlsx");

        Assert.Throws<FileNotFoundException>(() => RegionInventory.Inspect(missing));
    }

    [Fact]
    public void Inventory_command_writes_the_published_envelope()
    {
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-region-inventory-{Guid.NewGuid():N}.json");

        var exitCode = RegionInventory.Run([CreateWorkbook(R(1, ("A", "value", null))), output]);

        Assert.Equal(0, exitCode);
        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(output));
        Assert.Equal("tiwater.xlsx.region-inventory/v1", document.RootElement.GetProperty("schema").GetString());
        Assert.Equal(64, document.RootElement.GetProperty("inputSha256").GetString()!.Length);
    }

    [Fact]
    public void Inventory_v2_preserves_serial_and_display_text_while_publishing_1900_date_evidence()
    {
        var path = CreateDateWorkbook(uses1904Dates: false, serial: 46078D);

        var inventory = RegionInventory.InspectV2(path);

        Assert.Equal("tiwater.xlsx.region-inventory/v2", inventory.Schema);
        var cell = Assert.Single(Assert.Single(Assert.Single(inventory.Sheets).Regions).Rows).Cells.Single();
        Assert.Equal("46078", cell.RawValue);
        Assert.Equal("2026-02-25", cell.FormattedValue);
        Assert.Equal("date", cell.NormalizedValue.Kind);
        Assert.Equal("2026-02-25", cell.NormalizedValue.Iso8601);
    }

    [Fact]
    public void Inventory_v2_uses_declared_1904_date_system_for_a_different_serial_and_path()
    {
        var original = CreateDateWorkbook(uses1904Dates: true, serial: 3D);
        var relocated = Path.Combine(Path.GetTempPath(), $"relocated-region-date-{Guid.NewGuid():N}.xlsx");
        File.Copy(original, relocated);

        var left = RegionInventory.InspectV2(original);
        var right = RegionInventory.InspectV2(relocated);

        var cell = Assert.Single(Assert.Single(Assert.Single(left.Sheets).Regions).Rows).Cells.Single();
        var v1Cell = Assert.Single(Assert.Single(Assert.Single(RegionInventory.Inspect(original).Sheets).Regions).Rows).Cells.Single();
        Assert.Equal("3", cell.RawValue);
        Assert.Equal(v1Cell.RawValue, cell.RawValue);
        Assert.Equal(v1Cell.FormattedValue, cell.FormattedValue);
        Assert.Equal("1904-01-04", cell.NormalizedValue.Iso8601);
        Assert.Equal(left.InputSha256, right.InputSha256);
        Assert.Equal(
            System.Text.Json.JsonSerializer.Serialize(left.Sheets),
            System.Text.Json.JsonSerializer.Serialize(right.Sheets));
        Assert.NotEqual(left.File, right.File);
    }

    [Fact]
    public void Inventory_v2_and_published_evidence_share_custom_date_format_classification()
    {
        var path = CreateDateWorkbook(
            uses1904Dates: false,
            serial: 45588D,
            numberFormatId: 164U,
            customNumberFormat: "dd/mm/yyyy");

        var regionCell = Assert.Single(Assert.Single(Assert.Single(RegionInventory.InspectV2(path).Sheets).Regions).Rows).Cells.Single();
        using var evidenceDocument = System.Text.Json.JsonDocument.Parse(
            System.Text.Json.JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Json.Options));
        var evidenceCell = evidenceDocument.RootElement.GetProperty("sheets")[0].GetProperty("cells")[0];

        Assert.Equal("custom", evidenceCell.GetProperty("style").GetProperty("numberFormatEvidence").GetProperty("source").GetString());
        Assert.Equal("date", evidenceCell.GetProperty("style").GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal(
            evidenceCell.GetProperty("normalizedValue").GetProperty("kind").GetString(),
            regionCell.NormalizedValue.Kind);
        Assert.Equal(
            evidenceCell.GetProperty("normalizedValue").GetProperty("iso8601").GetString(),
            regionCell.NormalizedValue.Iso8601);
        Assert.Equal("2024-10-23", regionCell.NormalizedValue.Iso8601);
    }

    [Fact]
    public void Inventory_v1_shape_remains_unchanged_when_v2_is_available()
    {
        var path = CreateDateWorkbook(uses1904Dates: false, serial: 45292D);

        var json = System.Text.Json.JsonSerializer.Serialize(RegionInventory.Inspect(path), Json.Options);
        using var document = System.Text.Json.JsonDocument.Parse(json);
        var cell = document.RootElement.GetProperty("sheets")[0].GetProperty("regions")[0]
            .GetProperty("rows")[0].GetProperty("cells")[0];

        Assert.Equal("tiwater.xlsx.region-inventory/v1", document.RootElement.GetProperty("schema").GetString());
        Assert.False(cell.TryGetProperty("normalizedValue", out _));
        Assert.Equal(
            ["reference", "row", "column", "columnName", "rawValue", "formattedValue", "formula"],
            cell.EnumerateObject().Select(property => property.Name).ToArray());
    }

    [Fact]
    public void Inventory_command_requires_explicit_v2_opt_in_and_rejects_unknown_schema()
    {
        var path = CreateDateWorkbook(uses1904Dates: false, serial: 45292D);
        var defaultOutput = Path.Combine(Path.GetTempPath(), $"xlsx-region-v1-{Guid.NewGuid():N}.json");
        var v2Output = Path.Combine(Path.GetTempPath(), $"xlsx-region-v2-{Guid.NewGuid():N}.json");

        Assert.Equal(0, RegionInventory.Run([path, defaultOutput]));
        Assert.Equal(0, RegionInventory.Run([path, v2Output, "--schema", "v2"]));

        using var defaultDocument = System.Text.Json.JsonDocument.Parse(File.ReadAllText(defaultOutput));
        using var v2Document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(v2Output));
        Assert.Equal("tiwater.xlsx.region-inventory/v1", defaultDocument.RootElement.GetProperty("schema").GetString());
        Assert.Equal("tiwater.xlsx.region-inventory/v2", v2Document.RootElement.GetProperty("schema").GetString());
        Assert.Throws<InvalidOperationException>(() => RegionInventory.Run([path, "--schema", "v3"]));
    }

    [Fact]
    public void Inventory_path_change_does_not_change_workbook_facts()
    {
        var original = CreateWorkbook(R(3, ("D", "stable", null)));
        var copy = Path.Combine(Path.GetTempPath(), $"xlsx-region-copy-{Guid.NewGuid():N}.xlsx");
        File.Copy(original, copy);

        var left = RegionInventory.Inspect(original);
        var right = RegionInventory.Inspect(copy);

        Assert.Equal(left.InputSha256, right.InputSha256);
        Assert.Equal(
            System.Text.Json.JsonSerializer.Serialize(left.Sheets),
            System.Text.Json.JsonSerializer.Serialize(right.Sheets));
        Assert.NotEqual(left.File, right.File);
    }

    [Fact]
    public void Inventory_content_mutation_changes_attestation_and_preserves_duplicate_values()
    {
        var first = RegionInventory.Inspect(CreateWorkbook(R(1, ("A", "same", null), ("B", "same", null))));
        var second = RegionInventory.Inspect(CreateWorkbook(R(1, ("A", "same", null))));

        Assert.NotEqual(first.InputSha256, second.InputSha256);
        Assert.Equal(2, first.Sheets[0].Regions[0].Rows[0].Cells.Count);
        Assert.Single(second.Sheets[0].Regions[0].Rows[0].Cells);
    }

    [Fact]
    public void Inventory_formula_without_cached_value_is_material()
    {
        var path = CreateWorkbook(R(5, ("G", "", "1+1")));

        var cell = Assert.Single(Assert.Single(Assert.Single(RegionInventory.Inspect(path).Sheets).Regions).Rows).Cells.Single();

        Assert.Equal("G5", cell.Reference);
        Assert.Equal("1+1", cell.Formula);
        Assert.Equal(string.Empty, cell.RawValue);
    }

    [Fact]
    public void Inventory_rejects_non_workbook_content()
    {
        var path = Path.Combine(Path.GetTempPath(), $"not-workbook-{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(path, "not an Open XML package");

        Assert.ThrowsAny<Exception>(() => RegionInventory.Inspect(path));
    }

    private static string CreateWorkbook(params (uint Row, (string Column, string Value, string? Formula)[] Cells)[] rows)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-region-inventory-{Guid.NewGuid():N}.xlsx");
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        foreach (var sourceRow in rows)
        {
            var row = new Row { RowIndex = sourceRow.Row };
            foreach (var sourceCell in sourceRow.Cells)
            {
                var cell = new Cell
                {
                    CellReference = $"{sourceCell.Column}{sourceRow.Row}",
                    DataType = sourceCell.Formula is null ? CellValues.InlineString : null,
                    InlineString = sourceCell.Formula is null ? new InlineString(new Text(sourceCell.Value)) : null,
                    CellFormula = sourceCell.Formula is null ? null : new CellFormula(sourceCell.Formula),
                    CellValue = sourceCell.Formula is null ? null : new CellValue(sourceCell.Value)
                };
                row.Append(cell);
            }
            sheetData.Append(row);
        }
        worksheetPart.Worksheet = new Worksheet(sheetData);
        workbookPart.Workbook.AppendChild(new Sheets()).Append(
            new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Observed" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateDateWorkbook(
        bool uses1904Dates,
        double serial,
        uint numberFormatId = 14U,
        string? customNumberFormat = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-region-date-{Guid.NewGuid():N}.xlsx");
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new WorkbookProperties { Date1904 = uses1904Dates });
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()),
            new Fills(new Fill()),
            new Borders(new Border()),
            new CellStyleFormats(new CellFormat()),
            new CellFormats(
                new CellFormat(),
                new CellFormat { NumberFormatId = numberFormatId, ApplyNumberFormat = true }));
        if (customNumberFormat is not null)
        {
            stylesPart.Stylesheet.NumberingFormats = new NumberingFormats(
                new NumberingFormat { NumberFormatId = numberFormatId, FormatCode = customNumberFormat });
        }
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var cell = new Cell
        {
            CellReference = "F11",
            StyleIndex = 1U,
            CellValue = new CellValue(serial.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        };
        worksheetPart.Worksheet = new Worksheet(new SheetData(new Row(cell) { RowIndex = 11U }));
        workbookPart.Workbook.AppendChild(new Sheets()).Append(
            new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Calendar" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static (uint Row, (string Column, string Value, string? Formula)[] Cells) R(
        uint row,
        params (string Column, string Value, string? Formula)[] cells) => (row, cells);
}
