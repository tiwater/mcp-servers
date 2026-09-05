using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// Synthetic inputs only: failure must preserve bytes owned by a previous caller.
var root = Path.Combine(Path.GetTempPath(), $"fixed-command-preservation-{Guid.NewGuid():N}");
Directory.CreateDirectory(root);
var failures = new List<string>();
var cases = 0;
foreach (var format in new[] { "xlsx", "pptx" })
{
    Check(format, "existing-output", directory =>
    {
        var input = Path.Combine(directory, $"input.{format}");
        var output = Path.Combine(directory, $"output.{format}");
        Create(format, input);
        var original = new byte[] { 0, 255, 7, 19, 42 };
        File.WriteAllBytes(output, original);
        Require(Invoke(format, directory, input, output) == 1, "existing output must be rejected");
        Require(File.Exists(output) && File.ReadAllBytes(output).SequenceEqual(original), "existing output was deleted or changed");
    });
    Check(format, "existing-receipt", directory =>
    {
        var input = Path.Combine(directory, $"input.{format}");
        var output = Path.Combine(directory, $"output.{format}");
        Create(format, input);
        var before = File.ReadAllBytes(input);
        File.WriteAllText(Path.Combine(directory, "receipt.json"), "previous receipt");
        Require(Invoke(format, directory, input, output) == 1, "existing receipt must be rejected");
        Require(!File.Exists(output), "rejected call created output");
        Require(File.ReadAllBytes(input).SequenceEqual(before), "input changed");
        Require(File.ReadAllText(Path.Combine(directory, "receipt.json")) == "previous receipt", "receipt changed");
    });
    Check(format, "invalid-in-place", directory =>
    {
        var input = Path.Combine(directory, $"input.{format}");
        Create(format, input);
        var before = File.ReadAllBytes(input);
        Require(Invoke(format, directory, input, input, valid: false) == 1, "invalid target must be rejected");
        Require(File.ReadAllBytes(input).SequenceEqual(before), "failed in-place edit changed input");
        RequireNoTemporaryFiles(directory);
    });
    Check(format, "invalid-new-output", directory =>
    {
        var input = Path.Combine(directory, $"input.{format}");
        var output = Path.Combine(directory, $"output.{format}");
        Create(format, input);
        Require(Invoke(format, directory, input, output, valid: false) == 1, "invalid target must fail");
        Require(!File.Exists(output), "failed edit left output");
        RequireNoTemporaryFiles(directory);
    });
    foreach (var inPlace in new[] { false, true })
        Check(format, inPlace ? "valid-in-place" : "valid-new-output", directory =>
        {
            var input = Path.Combine(directory, $"input.{format}");
            var output = inPlace ? input : Path.Combine(directory, $"output.{format}");
            Create(format, input);
            var before = File.ReadAllBytes(input);
            Require(Invoke(format, directory, input, output) == 0, "valid edit rejected");
            if (!inPlace) Require(File.ReadAllBytes(input).SequenceEqual(before), "copy edit changed input");
            if (format == "xlsx")
            {
                using var document = SpreadsheetDocument.Open(output, false);
                var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>().Single(c => c.CellReference == "C4");
                Require(cell.CellValue?.Text == "73", "edited cell value differs");
            }
            else
            {
                using var document = PresentationDocument.Open(output, false);
                var transform = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Transform2D>().Single();
                Require(transform.Offset!.X == 731 && transform.Extents!.Cx == 2900, "edited geometry differs");
            }
            using var receipt = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "receipt.json")));
            Require(receipt.RootElement.GetProperty("pass").GetBoolean(), "receipt not passing");
            RequireNoTemporaryFiles(directory);
        });
}
Console.WriteLine(JsonSerializer.Serialize(new { cases, failures, artifacts = root }));
return failures.Count == 0 ? 0 : 1;

void Check(string format, string name, Action<string> action)
{
    cases++;
    var directory = Path.Combine(root, $"{format}-{name}");
    Directory.CreateDirectory(directory);
    try { action(directory); }
    catch (Exception error) { failures.Add($"{format}/{name}: {error.Message}"); }
}

static void Require(bool value, string message)
{
    if (!value) throw new InvalidOperationException(message);
}

static void RequireNoTemporaryFiles(string directory)
    => Require(!Directory.EnumerateFiles(directory).Any(file => Path.GetFileName(file).StartsWith('.')), "temporary files leaked");

static int Invoke(string format, string directory, string input, string output, bool valid = true)
{
    object changes = format == "xlsx"
        ? new[] { new { sheet = valid ? "Synthetic" : "Absent", cell = "C4", value = 73 } }
        : new[] { new { slideNumber = 1, shapeId = valid ? 7 : 999, x = 731, y = 952, cx = 2900, cy = 3800 } };
    var request = Path.Combine(directory, "request.json");
    File.WriteAllText(request, JsonSerializer.Serialize(new { input, output, receiptOutput = Path.Combine(directory, "receipt.json"), changes }));
    var savedOut = Console.Out;
    var savedError = Console.Error;
    using var capture = new StringWriter();
    try
    {
        Console.SetOut(capture);
        Console.SetError(capture);
        return format == "xlsx"
            ? Dockit.Xlsx.FixedCommandRunner.Run("xlsx_set_cell_value", [request])
            : Dockit.Pptx.FixedCommandRunner.Run("pptx_set_shape_geometry", [request]);
    }
    finally
    {
        Console.SetOut(savedOut);
        Console.SetError(savedError);
        File.WriteAllText(Path.Combine(directory, "command.log"), capture.ToString());
    }
}

static void Create(string format, string file)
{
    if (format == "xlsx")
    {
        using var document = SpreadsheetDocument.Create(file, SpreadsheetDocumentType.Workbook);
        var workbook = document.AddWorkbookPart();
        var sheet = workbook.AddNewPart<WorksheetPart>();
        sheet.Worksheet = new S.Worksheet(new S.SheetData(new S.Row(new S.Cell { CellReference = "C4", CellValue = new S.CellValue("11") }) { RowIndex = 4 }));
        workbook.Workbook = new S.Workbook(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Synthetic" }));
    }
    else
    {
        using var document = PresentationDocument.Create(file, PresentationDocumentType.Presentation);
        var presentation = document.AddPresentationPart();
        var slide = presentation.AddNewPart<SlidePart>();
        slide.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = "Root" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(),
            new P.Shape(new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = 7, Name = "Synthetic" }, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 100, Y = 200 }, new A.Extents { Cx = 3000, Cy = 4000 }))))));
        presentation.Presentation = new P.Presentation(new P.SlideIdList(new P.SlideId { Id = 256, RelationshipId = presentation.GetIdOfPart(slide) }));
    }
}
