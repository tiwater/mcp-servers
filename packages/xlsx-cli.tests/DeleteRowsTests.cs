using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Dockit.Xlsx.Tests;

public class DeleteRowsTests
{
    [Fact]
    public void Contract_declares_delete_rows_and_typed_failure_evidence()
    {
        var contractRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../xlsx-cli/contracts"));
        var request = File.ReadAllText(Path.Combine(contractRoot, "tiwater.xlsx-edit-v1.schema.json"));
        var result = File.ReadAllText(Path.Combine(contractRoot, "tiwater.xlsx-edit-result-v1.schema.json"));

        Assert.Contains("\"deleteRows\"", request, StringComparison.Ordinal);
        Assert.Contains("\"errorCode\"", result, StringComparison.Ordinal);
        Assert.DoesNotContain("errorCode", JsonSerializer.Serialize(new XlsxEditAppliedOperation("insertRows", true, "ok"), Json.Options), StringComparison.Ordinal);
        Assert.Contains("xlsx.deleteRows.test", JsonSerializer.Serialize(new XlsxEditAppliedOperation("deleteRows", false, "failed", ErrorCode: "xlsx.deleteRows.test"), Json.Options), StringComparison.Ordinal);
    }

    [Fact]
    public void Delete_rows_is_structural_and_translates_surviving_content()
    {
        var input = CreateWorkbook(10);
        var output = TemporaryWorkbook("positive");

        var result = Editor.Apply(input, output, [
            new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: 3, Count: 2)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);
        Assert.Null(operation.ErrorCode);
        Assert.Equal("3:4", operation.ChangedRange);

        using var workbook = SpreadsheetDocument.Open(output, false);
        var data = Worksheet(workbook.WorkbookPart!, "Data");
        var other = Worksheet(workbook.WorkbookPart!, "Other");
        var rows = data.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        Assert.Equal(8, rows.Count);
        Assert.DoesNotContain(rows, row => row.RowIndex?.Value > 8);
        Assert.Equal("row-5", CellText(data, "A3"));
        Assert.Equal(7U, Cell(data, "A3").StyleIndex?.Value);
        Assert.Equal("A6+1", Cell(data, "B3").CellFormula?.Text);
        Assert.Equal("Data!A6", Cell(other, "A1").CellFormula?.Text);
        Assert.Equal("A1:C8", data.GetFirstChild<SheetDimension>()?.Reference?.Value);
        Assert.Contains(data.Descendants<MergeCell>(), merge => merge.Reference?.Value == "C3:D5");
        var dataPart = (WorksheetPart)workbook.WorkbookPart!.GetPartById(workbook.WorkbookPart.Workbook.Sheets!.Elements<Sheet>().First().Id!.Value!);
        Assert.Equal("D4", Assert.Single(dataPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>()).Reference?.Value);
        Assert.Equal("3", Assert.Single(dataPart.DrawingsPart!.WorksheetDrawing.Descendants<Xdr.RowId>()).Text);

        var names = workbook.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.Contains(names, name => name.Name?.Value == "_xlnm.Print_Area" && name.Text == "'Data'!$A$1:$D$8");
        Assert.Contains(names, name => name.Name?.Value == "_xlnm.Print_Titles" && name.Text == "'Data'!$1:$1,'Data'!$A:$B");
        var breaks = data.GetFirstChild<RowBreaks>()!;
        Assert.Equal([4U], breaks.Elements<Break>().Select(item => item.Id!.Value).ToArray());
        Assert.Equal(OrientationValues.Landscape, data.GetFirstChild<PageSetup>()?.Orientation?.Value);
        Assert.Equal(1U, data.GetFirstChild<PageSetup>()?.FitToWidth?.Value);
        Assert.True(Validator.Validate(output).Valid);
    }

    [Fact]
    public void Delete_rows_rejects_a_surviving_formula_that_targets_deleted_cells()
    {
        var input = CreateWorkbook(7);
        var output = TemporaryWorkbook("formula-failure");
        using (var workbook = SpreadsheetDocument.Open(input, true))
        {
            Cell(Worksheet(workbook.WorkbookPart!, "Other"), "A1").CellFormula = new CellFormula("Data!A3");
        }

        var result = Editor.Apply(input, output, [
            new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: 3, Count: 1)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.False(operation.Applied);
        Assert.Equal("xlsx.deleteRows.formulaTargetsDeletedRows", operation.ErrorCode);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Delete_rows_fails_closed_for_unsupported_range_metadata()
    {
        var input = CreateWorkbook(7, includeFormulas: false);
        var output = TemporaryWorkbook("metadata-failure");
        using (var workbook = SpreadsheetDocument.Open(input, true))
        {
            var worksheet = Worksheet(workbook.WorkbookPart!, "Data");
            worksheet.InsertAfter(new AutoFilter { Reference = "A1:A7" }, worksheet.GetFirstChild<SheetData>());
            worksheet.Save();
        }

        var result = Editor.Apply(input, output, [new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: 3, Count: 1)]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.False(operation.Applied);
        Assert.Equal("xlsx.deleteRows.unsupportedDependentStructure", operation.ErrorCode);
        Assert.False(File.Exists(output));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(1048576, 2)]
    public void Delete_rows_rejects_unbounded_coordinates(int startRow, int count)
    {
        var input = CreateWorkbook(3);
        var output = TemporaryWorkbook("bounds");

        var result = Editor.Apply(input, output, [
            new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: startRow, Count: count)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.False(operation.Applied);
        Assert.Equal("xlsx.edit.invalidCoordinates", operation.ErrorCode);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Delete_rows_failure_is_atomic_for_an_existing_output()
    {
        var input = CreateWorkbook(6);
        var output = TemporaryWorkbook("atomic");
        var sentinel = new byte[] { 2, 3, 5, 7 };
        File.WriteAllBytes(output, sentinel);
        using (var workbook = SpreadsheetDocument.Open(input, true))
        {
            Cell(Worksheet(workbook.WorkbookPart!, "Other"), "A1").CellFormula = new CellFormula("Data!A2");
        }

        var result = Editor.Apply(input, output, [
            new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: 2, Count: 1)
        ]);

        Assert.False(Assert.Single(result.AppliedOperations).Applied);
        Assert.Equal(sentinel, File.ReadAllBytes(output));
    }

    [Fact]
    public void Delete_rows_property_preserves_order_for_unseen_shapes()
    {
        for (var rowCount = 5; rowCount <= 12; rowCount++)
        {
            for (var startRow = 2; startRow < rowCount; startRow++)
            {
                var count = Math.Min(2, rowCount - startRow + 1);
                var input = CreateWorkbook(rowCount, includeFormulas: false);
                var output = TemporaryWorkbook("property");
                var result = Editor.Apply(input, output, [new XlsxEditOperation("deleteRows", Sheet: "Data", StartRow: startRow, Count: count)]);
                Assert.True(Assert.Single(result.AppliedOperations).Applied);

                using var workbook = SpreadsheetDocument.Open(output, false);
                var worksheet = Worksheet(workbook.WorkbookPart!, "Data");
                var actual = worksheet.GetFirstChild<SheetData>()!.Elements<Row>()
                    .Select(row => CellText(worksheet, $"A{row.RowIndex!.Value}"))
                    .ToArray();
                var expected = Enumerable.Range(1, rowCount)
                    .Where(row => row < startRow || row >= startRow + count)
                    .Select(row => $"row-{row}")
                    .ToArray();
                Assert.Equal(expected, actual);
            }
        }
    }

    [Fact]
    public void Existing_insert_rows_contract_remains_compatible()
    {
        var input = CreateWorkbook(5, includeFormulas: false);
        var output = TemporaryWorkbook("compatibility");
        var result = Editor.Apply(input, output, [new XlsxEditOperation("insertRows", Sheet: "Data", StartRow: 3, Count: 1)]);
        Assert.True(Assert.Single(result.AppliedOperations).Applied);
        using var workbook = SpreadsheetDocument.Open(output, false);
        Assert.Equal("row-3", CellText(Worksheet(workbook.WorkbookPart!, "Data"), "A4"));
    }

    private static string CreateWorkbook(int rowCount, bool includeFormulas = true)
    {
        var path = TemporaryWorkbook("fixture");
        using var workbook = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = workbook.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(Enumerable.Range(0, 8).Select(_ => new Font())) { Count = 8 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(Enumerable.Range(0, 8).Select(index => new CellFormat { FontId = (uint)index })) { Count = 8 });

        var dataPart = workbookPart.AddNewPart<WorksheetPart>();
        var rows = new List<Row>();
        for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
        {
            var row = new Row { RowIndex = (uint)rowIndex };
            row.Append(new Cell { CellReference = $"A{rowIndex}", DataType = CellValues.InlineString, InlineString = new InlineString(new Text($"row-{rowIndex}")), StyleIndex = (uint)(rowIndex == 5 ? 7 : 1) });
            if (includeFormulas && rowIndex == 5)
                row.Append(new Cell { CellReference = "B5", CellFormula = new CellFormula("A8+1"), CellValue = new CellValue("9") });
            rows.Add(row);
        }
        dataPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = $"A1:C{rowCount}" },
            new SheetData(rows),
            rowCount >= 7 ? new MergeCells(new MergeCell { Reference = "C5:D7" }) : new MergeCells(),
            new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D },
            new PageSetup { Orientation = OrientationValues.Landscape, FitToWidth = 1U },
            new RowBreaks(new Break { Id = 3U, Max = 16383U, ManualPageBreak = true }, new Break { Id = 6U, Max = 16383U, ManualPageBreak = true }) { Count = 2U, ManualBreakCount = 2U });

        if (includeFormulas)
        {
            var commentsPart = dataPart.AddNewPart<WorksheetCommentsPart>();
            commentsPart.Comments = new Comments(
                new Authors(new Author("tester")),
                new CommentList(
                    new Comment(new CommentText(new Run(new Text("deleted")))) { Reference = "D3", AuthorId = 0U },
                    new Comment(new CommentText(new Run(new Text("shifted")))) { Reference = "D6", AuthorId = 0U }));

            var drawingsPart = dataPart.AddNewPart<DrawingsPart>();
            var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (var image = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var bytes = System.Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=");
                image.Write(bytes);
            }
            var picture = new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 1U, Name = "Picture 1" },
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                new Xdr.BlipFill(
                    new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 9525L, Cy = 9525L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("5"), new Xdr.RowOffset("0")),
                    new Xdr.Extent { Cx = 9525L, Cy = 9525L },
                    picture,
                    new Xdr.ClientData()));
            dataPart.Worksheet.Append(new Drawing { Id = dataPart.GetIdOfPart(drawingsPart) });
        }

        var otherPart = workbookPart.AddNewPart<WorksheetPart>();
        var otherCell = includeFormulas
            ? new Cell { CellReference = "A1", CellFormula = new CellFormula("Data!A8"), CellValue = new CellValue("8") }
            : new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("other")) };
        otherPart.Worksheet = new Worksheet(new SheetData(new Row(otherCell) { RowIndex = 1 }));
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(dataPart), SheetId = 1, Name = "Data" });
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(otherPart), SheetId = 2, Name = "Other" });
        workbookPart.Workbook.DefinedNames = new DefinedNames(
            new DefinedName($"'Data'!$A$1:$D${rowCount}") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
            new DefinedName("'Data'!$1:$1,'Data'!$A:$B") { Name = "_xlnm.Print_Titles", LocalSheetId = 0 });
        workbookPart.Workbook.Save();
        dataPart.Worksheet.Save();
        otherPart.Worksheet.Save();
        return path;
    }

    private static string TemporaryWorkbook(string label) => Path.Combine(Path.GetTempPath(), $"xlsx-delete-rows-{label}-{Guid.NewGuid():N}.xlsx");

    private static Worksheet Worksheet(WorkbookPart workbookPart, string name)
    {
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name?.Value == name);
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet;
    }

    private static Cell Cell(Worksheet worksheet, string reference) => worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);

    private static string CellText(Worksheet worksheet, string reference) => Cell(worksheet, reference).InlineString?.InnerText ?? Cell(worksheet, reference).CellValue?.Text ?? string.Empty;
}
