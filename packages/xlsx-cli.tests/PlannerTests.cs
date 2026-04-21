using Dockit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class PlannerTests
{
    [Fact]
    public void Inspect_reports_used_range_and_formula_cells_without_note_inference()
    {
        var path = CreateAna14WorkbookFixture();

        var report = Inspector.Inspect(path);
        var sheet = Assert.Single(report.Sheets);

        Assert.Equal("D5:L15", sheet.UsedRange);
        Assert.True(sheet.FormulaCellCount >= 5);
        Assert.Empty(sheet.NoteRows ?? []);
    }

    [Fact]
    public void Plan_selects_summarized_tables_and_emits_fixed_layout_edits()
    {
        var path = CreateAna14WorkbookFixture();
        var request = new XlsxPlanRequest(
            Scenario: "experimental-record-attachment",
            Sheet: "Sheet1",
            Sources:
            [
                new XlsxPlanSourceDocument(
                    "attachment-280",
                    "sample-280.pdf",
                    [
                        new XlsxPlanSourceTable(
                            "Area Summarized by Name",
                            ["", "SampleName", "Result Id", "LC", "LC 1d _", "HC", "HC 1d _", "HC 2d _", "HC 3d _", "post HC-3d"],
                            [
                                ["", "SampleName", "Result Id", "LC", "LC 1d _", "HC", "HC 1d _", "HC 2d _", "HC 3d _", "post HC-3d"],
                                ["1", "260359-01", "3936", "233988", "383789", "394821", "525522", "787624", "476758", "52470"],
                                ["2", "HSPXXXXSTD01", "3937", "252353", "341366", "516287", "397165", "711337", "521517", "69410"],
                            ],
                            Page: 5)
                    ]),
                new XlsxPlanSourceDocument(
                    "attachment-360",
                    "sample-360.pdf",
                    [
                        new XlsxPlanSourceTable(
                            "Area Summarized by Name",
                            ["", "SampleName", "Result Id", "LC", "LC 1d _", "HC", "HC 1d _", "HC 2d _", "HC 3d _", "post HC-3d"],
                            [
                                ["", "SampleName", "Result Id", "LC", "LC 1d _", "HC", "HC 1d _", "HC 2d _", "HC 3d _", "post HC-3d"],
                                ["1", "260359-01", "3935", "0", "139924", "0", "72690", "158882", "117073", "12161"],
                                ["2", "HSPXXXXSTD01", "3938", "0", "125292", "0", "53132", "142995", "126642", "14593"],
                            ],
                            Page: 5)
                    ]),
            ]);

        var result = Planner.Plan(path, request);

        Assert.Equal("high", result.Confidence);
        Assert.Empty(result.Warnings);
        Assert.Equal(2, result.SelectedSources.Count);
        Assert.Equal(3, result.Sections.Count);
        Assert.Equal(2, result.ProposedEdits.Count);

        var firstEdit = result.ProposedEdits[0];
        Assert.Equal("setRangeValues", firstEdit.Type);
        Assert.Equal("Sheet1", firstEdit.Sheet);
        Assert.Equal("E6", firstEdit.StartCell);
        Assert.Equal("52470", firstEdit.Values![0][6]);

        var secondEdit = result.ProposedEdits[1];
        Assert.Equal("E9", secondEdit.StartCell);
        Assert.Equal("117073", secondEdit.Values![0][5]);

        var formulaSection = Assert.Single(result.Sections.Where(section => section.FormulaDriven));
        Assert.Equal("E12", formulaSection.TargetStartCell);
        Assert.Equal("K13", formulaSection.TargetEndCell);
    }

    private static string CreateAna14WorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana14-workbook-{Guid.NewGuid():N}.xlsx");

        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new SharedStringTable();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            CreateRow(5, ("D5", "280 nm峰面积"), ("E5", "LC"), ("F5", "LC_1d"), ("G5", "HC"), ("H5", "HC_1d"), ("I5", "HC_2d"), ("J5", "HC_3d"), ("K5", "post HC_3d"), ("L5", "DAR")),
            CreateNumericRow(6, [("D6", "260359-01"), ("E6", "233988"), ("F6", "383789"), ("G6", "394821"), ("H6", "525522"), ("I6", "787624"), ("J6", "476758"), ("K6", "52470"), ("L6", "4.53987553812785")], ["L6"]),
            CreateNumericRow(7, [("D7", "HSPXXXXSTD01"), ("E7", "252353"), ("F7", "341366"), ("G7", "516287"), ("H7", "397165"), ("I7", "711337"), ("J7", "521517"), ("K7", "69410"), ("L7", "4.39277688253888")], ["L7"]),
            CreateRow(8, ("D8", "360 nm峰面积"), ("E8", "LC"), ("F8", "LC_1d"), ("G8", "HC"), ("H8", "HC_1d"), ("I8", "HC_2d"), ("J8", "HC_3d"), ("K8", "post HC_3d")),
            CreateNumericRow(9, [("D9", "260359-01"), ("E9", "0"), ("F9", "139924"), ("G9", "0"), ("H9", "72690"), ("I9", "158882"), ("J9", "117073"), ("K9", "12161")]),
            CreateNumericRow(10, [("D10", "HSPXXXXSTD01"), ("E10", "0"), ("F10", "125292"), ("G10", "0"), ("H10", "53132"), ("I10", "142995"), ("J10", "126642"), ("K10", "14593")]),
            CreateRow(11, ("D11", "280 nm峰面积-360 nm峰面积*0.784"), ("E11", "LC"), ("F11", "LC_1d"), ("G11", "HC"), ("H11", "HC_1d"), ("I11", "HC_2d"), ("J11", "HC_3d"), ("K11", "post HC_3d"), ("L11", "DAR")),
            CreateNumericRow(12, [("D12", "260359-01"), ("E12", "233988"), ("F12", "274088.584"), ("G12", "394821"), ("H12", "468533.04"), ("I12", "663060.512"), ("J12", "384972.768"), ("K12", "42935.776"), ("L12", "4.22925457695732")], ["E12", "F12", "G12", "H12", "I12", "J12", "K12", "L12"]),
            CreateNumericRow(13, [("D13", "HSPXXXXSTD01"), ("E13", "252353"), ("F13", "243137.072"), ("G13", "516287"), ("H13", "355509.512"), ("I13", "599228.92"), ("J13", "422229.672"), ("K13", "57969.088"), ("L13", "4.05082073919091")], ["E13", "F13", "G13", "H13", "I13", "J13", "K13", "L13"]),
            CreateRow(15, ("D15", "注意：从两个PDF中的 Area Summarized by Name 表按样品编号取值，并保留第三部分的公式计算。"))
        );
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "D5:L15" },
            sheetData,
            new MergeCells(new MergeCell { Reference = "D15:L15" }));
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static Row CreateRow(uint rowIndex, params (string Ref, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (cellRef, value) in cells)
        {
            row.Append(new Cell
            {
                CellReference = cellRef,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value))
            });
        }

        return row;
    }

    private static Row CreateNumericRow(uint rowIndex, (string Ref, string Value)[] cells, string[]? formulaCellRefs = null)
    {
        var formulaRefs = new HashSet<string>(formulaCellRefs ?? [], StringComparer.OrdinalIgnoreCase);
        var row = new Row { RowIndex = rowIndex };
        foreach (var (cellRef, value) in cells)
        {
            var cell = new Cell { CellReference = cellRef };
            if (cellRef.StartsWith("D", StringComparison.Ordinal))
            {
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(value));
            }
            else
            {
                cell.CellValue = new CellValue(value);
            }

            if (formulaRefs.Contains(cellRef))
            {
                cell.CellFormula = new CellFormula("1+1");
            }

            row.Append(cell);
        }

        return row;
    }
}
