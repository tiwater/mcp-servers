using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class TableHeaderRepeatTests
{
    [Fact]
    public void Repeat_header_batch_sets_and_unsets_body_header_and_footer_rows()
    {
        var input = CreateStoryTableFixture();
        var output = TempDocx("repeat-stories");

        var result = Editor.Apply(input, output, [
            new DocxEditOperation("setTableRowRepeatAsHeader", TableIndex: 0, RowIndex: 0, RepeatAsHeader: true),
            new DocxEditOperation("setTableRowRepeatAsHeader", HeaderIndex: 0, TableIndex: 0, RowIndex: 0, RepeatAsHeader: true),
            new DocxEditOperation("setTableRowRepeatAsHeader", FooterIndex: 0, TableIndex: 0, RowIndex: 0, RepeatAsHeader: false),
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        var inspection = Inspector.InspectTables(output);
        Assert.Equal("tiwater.docx.inspect-tables/v1", inspection.Schema);
        Assert.Equal("body-and-nested-depth-first", inspection.ExtractionView["tableTraversal"]);
        Assert.Equal("header-footer-and-nested-depth-first", inspection.ExtractionView["storyTableTraversal"]);

        var body = Assert.Single(inspection.Tables, table => table.MutationAddress is { Kind: "body", TableIndex: 0 });
        Assert.True(body.Rows[0].RepeatAsHeader);
        Assert.Equal(2, body.Rows[0].Cells[0].GridSpan);
        Assert.Equal("restart", body.Rows[1].Cells[0].VMerge);

        var header = Assert.Single(inspection.StoryTables!, table => table.MutationAddress is { Kind: "header", HeaderIndex: 0, TableIndex: 0 });
        Assert.True(header.Rows[0].RepeatAsHeader);
        Assert.Equal(2, header.Story!.References.Count);
        Assert.All(header.Story.References, reference => Assert.Equal(0, reference.SectionIndex));
        Assert.Equal(["default", "even"], header.Story.References.Select(reference => reference.ReferenceType).ToArray());

        var footer = Assert.Single(inspection.StoryTables!, table => table.MutationAddress is { Kind: "footer", FooterIndex: 0, TableIndex: 0 });
        Assert.False(footer.Rows[0].RepeatAsHeader);
        var footerReference = Assert.Single(footer.Story!.References);
        Assert.Equal(0, footerReference.SectionIndex);
        Assert.Equal("first", footerReference.ReferenceType);

        using var edited = WordprocessingDocument.Open(output, false);
        var validationErrors = new OpenXmlValidator().Validate(edited).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors.Select(error => error.Description)));
    }

    [Fact]
    public void Inspection_observes_nested_and_unseen_story_topology_without_making_nested_tables_mutable()
    {
        var inspection = Inspector.InspectTables(CreateStoryTableFixture());

        Assert.Single(inspection.Tables);
        Assert.Equal(3, inspection.StoryTables!.Count);
        var nested = Assert.Single(inspection.StoryTables, table => table.ParentCellAddress is not null);
        Assert.Equal("header", nested.Story!.Kind);
        Assert.Null(nested.MutationAddress);
        Assert.Equal("nested header value", nested.Rows[0].Cells[0].Text);
        Assert.Contains("cell:0", nested.ContainmentPath);

        var header = Assert.Single(inspection.StoryTables, table => table.MutationAddress?.Kind == "header");
        Assert.Equal(1, header.Rows[0].GridBefore);
        Assert.Equal(1, header.Rows[0].GridAfter);
        Assert.Equal(4, header.Rows[0].GridWidth);
        Assert.Equal("header direct", header.Rows[0].Cells[0].Text);

        var json = JsonSerializer.Serialize(inspection, Json.Options);
        using var parsed = JsonDocument.Parse(json);
        Assert.Equal("tiwater.docx.inspect-tables/v1", parsed.RootElement.GetProperty("Schema").GetString());
        Assert.True(parsed.RootElement.GetProperty("Tables")[0].TryGetProperty("TableIndex", out _));
        Assert.True(parsed.RootElement.GetProperty("Tables")[0].GetProperty("Rows")[0].GetProperty("Cells")[0].TryGetProperty("Text", out _));
    }

    [Fact]
    public void Repeat_header_batch_rejects_missing_ambiguous_invalid_duplicate_and_nested_targets_atomically()
    {
        var input = CreateStoryTableFixture();
        var cases = new IReadOnlyList<DocxEditOperation>[]
        {
            [
                new("setTableRowRepeatAsHeader", TableIndex: 0, RowIndex: 0, RepeatAsHeader: true),
                new("setTableRowRepeatAsHeader", HeaderIndex: 9, TableIndex: 0, RowIndex: 0, RepeatAsHeader: true),
            ],
            [new("setTableRowRepeatAsHeader", HeaderIndex: 0, FooterIndex: 0, TableIndex: 0, RowIndex: 0, RepeatAsHeader: true)],
            [new("setTableRowRepeatAsHeader", TableIndex: -1, RowIndex: 0, RepeatAsHeader: true)],
            [new("setTableRowRepeatAsHeader", TableIndex: 0, RowIndex: 0)],
            [
                new("setTableRowRepeatAsHeader", TableIndex: 0, RowIndex: 0, RepeatAsHeader: true),
                new("setTableRowRepeatAsHeader", TableIndex: 0, RowIndex: 0, RepeatAsHeader: false),
            ],
            [new("setTableRowRepeatAsHeader", HeaderIndex: 0, TableIndex: 1, RowIndex: 0, RepeatAsHeader: true)],
        };

        foreach (var operations in cases)
        {
            var output = TempDocx("repeat-rejected");
            var result = Editor.Apply(input, output, operations);
            Assert.All(result.AppliedOperations, operation => Assert.False(operation.Applied));
            var body = Assert.Single(Inspector.InspectTables(output).Tables, table => table.MutationAddress is { Kind: "body", TableIndex: 0 });
            Assert.False(body.Rows[0].RepeatAsHeader);
        }
    }

    private static string CreateStoryTableFixture()
    {
        var path = TempDocx("repeat-fixture");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();

        var header = main.AddNewPart<HeaderPart>();
        var nested = TableWithText("nested header value");
        var headerCell = new TableCell(new Paragraph(new Run(new Text("header direct"))), nested);
        headerCell.TableCellProperties = new TableCellProperties(new GridSpan { Val = 2 });
        var headerRow = new TableRow(
            new TableRowProperties(new GridBefore { Val = 1 }, new GridAfter { Val = 1 }),
            headerCell);
        header.Header = new Header(new Table(new TableProperties(), new TableGrid(new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn()), headerRow));
        header.Header.Save();

        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Table(new TableProperties(), new TableGrid(new GridColumn()), new TableRow(
            new TableRowProperties(new TableHeader()),
            new TableCell(new Paragraph(new Run(new Text("footer value")))))));
        footer.Footer.Save();

        var bodyTable = new Table(new TableProperties(), new TableGrid(new GridColumn(), new GridColumn()),
            new TableRow(new TableCell(
                new TableCellProperties(new GridSpan { Val = 2 }),
                new Paragraph(new Run(new Text("body merged"))))),
            new TableRow(new TableCell(
                new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                new Paragraph(new Run(new Text("body vertical"))))));
        var section = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
            new HeaderReference { Type = HeaderFooterValues.Even, Id = main.GetIdOfPart(header) },
            new FooterReference { Type = HeaderFooterValues.First, Id = main.GetIdOfPart(footer) });
        main.Document = new Document(new Body(bodyTable, section));
        main.Document.Save();
        return path;
    }

    private static Table TableWithText(string text) =>
        new(new TableProperties(), new TableGrid(new GridColumn()), new TableRow(new TableCell(new Paragraph(new Run(new Text(text))))));

    private static string TempDocx(string prefix) =>
        Path.Combine(Path.GetTempPath(), $"{prefix}-{Guid.NewGuid():N}.docx");
}
