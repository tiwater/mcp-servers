using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class TableRangeCopy
{
    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException("copy-table-range requires <request.json>");
        var request = JsonSerializer.Deserialize<DocxCopyTableRangeRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("copy-table-range-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
        return 0;
    }

    public static DocxCopyTableRangeReceipt Apply(DocxCopyTableRangeRequest request)
    {
        if (request.Schema != "tiwater.docx-copy-table-range/v1")
            throw new InvalidOperationException("copy-table-range-schema-unsupported");
        if (request.Source.RowRefs.Count == 0 || request.Target.RowPatternRefs.Count == 0 || request.Columns.Count == 0)
            throw new InvalidOperationException("source rows, target row pattern, and columns must be non-empty");

        var sourcePath = Path.GetFullPath(request.Source.Input);
        var targetPath = Path.GetFullPath(request.Target.Input);
        var outputPath = Path.GetFullPath(request.Output);
        var receiptPath = Path.GetFullPath(request.ReceiptOutput);
        RequireNewPath(outputPath, "output");
        RequireNewPath(receiptPath, "receiptOutput");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, receiptPath))
            throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, sourcePath)
            || StringComparer.OrdinalIgnoreCase.Equals(outputPath, targetPath))
            throw new InvalidOperationException("output-must-not-overwrite-an-input");

        var sourceRefs = Observation.ResolveReferences(sourcePath, request.Source.Revision, request.Source.RowRefs);
        var targetRefs = Observation.ResolveReferences(
            targetPath,
            request.Target.Revision,
            [request.Target.TableRef, .. request.Target.RowPatternRefs]);
        if (sourceRefs.Any(item => item.Kind != "row")) throw new InvalidOperationException("source-ref-wrong-kind");
        if (targetRefs[0].Kind != "table" || targetRefs.Skip(1).Any(item => item.Kind != "row"))
            throw new InvalidOperationException("target-ref-wrong-kind");
        if (request.Columns.Select(item => item.SourceGridColumn).Distinct().Count() != request.Columns.Count
            || request.Columns.Select(item => item.TargetGridColumn).Distinct().Count() != request.Columns.Count)
            throw new InvalidOperationException("column-mapping-must-be-one-to-one");
        if (request.Columns.Any(item => item.SourceGridColumn < 0 || item.TargetGridColumn < 0))
            throw new InvalidOperationException("grid-columns-must-be-nonnegative");

        IReadOnlyList<IReadOnlyList<string>> sourceValues;
        IReadOnlyList<TableRow> patternRows;
        Table targetTable;
        int targetStartRowIndex;
        IReadOnlyDictionary<string, int> baselineValidationIssues;
        using (var sourceDocument = WordprocessingDocument.Open(sourcePath, false))
        using (var targetDocument = WordprocessingDocument.Open(targetPath, false))
        {
            var sourceRows = sourceRefs
                .Select(item => Observation.ResolveNativePath(sourceDocument, item.StoryPart, item.NativePath))
                .OfType<TableRow>().ToList();
            EnsureContiguousRows(sourceRows, "source");

            targetTable = Observation.ResolveNativePath(targetDocument, targetRefs[0].StoryPart, targetRefs[0].NativePath) as Table
                ?? throw new InvalidOperationException("target-table-not-found");
            patternRows = targetRefs.Skip(1)
                .Select(item => Observation.ResolveNativePath(targetDocument, item.StoryPart, item.NativePath))
                .OfType<TableRow>().ToList();
            EnsureContiguousRows(patternRows, "target-pattern");
            if (patternRows.Any(row => !ReferenceEquals(row.Parent, targetTable)))
                throw new InvalidOperationException("target-pattern-rows-must-belong-to-target-table");
            targetStartRowIndex = targetTable.Elements<TableRow>().ToList().IndexOf(patternRows[0]);
            if (targetStartRowIndex < 0) throw new InvalidOperationException("target-pattern-start-not-found");
            if (HasVerticalMerge(patternRows) && sourceRows.Count % patternRows.Count != 0)
                throw new InvalidOperationException("vertical-merge-pattern-must-repeat-completely");

            var sourceCells = new HashSet<TableCell>();
            sourceValues = sourceRows.Select(row => (IReadOnlyList<string>)request.Columns.Select(mapping =>
            {
                var cell = CellAtGridColumn(row, mapping.SourceGridColumn)
                    ?? throw new InvalidOperationException($"source-grid-column-out-of-range: {mapping.SourceGridColumn}");
                if (!sourceCells.Add(cell)) throw new InvalidOperationException("column-mapping-selects-the-same-source-cell-more-than-once");
                EnsureTextOnly(cell, "source");
                return string.Join("\n", cell.Elements<Paragraph>().Select(paragraph => paragraph.InnerText));
            }).ToList()).ToList();

            foreach (var patternRow in patternRows)
            {
                var targetCells = new HashSet<TableCell>();
                foreach (var mapping in request.Columns)
                {
                    var cell = CellAtGridColumn(patternRow, mapping.TargetGridColumn)
                        ?? throw new InvalidOperationException($"target-grid-column-out-of-range: {mapping.TargetGridColumn}");
                    if (!targetCells.Add(cell)) throw new InvalidOperationException("column-mapping-selects-the-same-target-cell-more-than-once");
                    EnsureTextOnly(cell, "target");
                }
            }
            baselineValidationIssues = ValidationIssueCounts(targetDocument);
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(targetPath, temporaryPath, overwrite: false);
            using (var outputDocument = WordprocessingDocument.Open(temporaryPath, true))
            {
                var outputTable = Observation.ResolveNativePath(outputDocument, targetRefs[0].StoryPart, targetRefs[0].NativePath) as Table
                    ?? throw new InvalidOperationException("target-table-not-found-in-output");
                var outputPatterns = targetRefs.Skip(1)
                    .Select(item => Observation.ResolveNativePath(outputDocument, item.StoryPart, item.NativePath))
                    .OfType<TableRow>().ToList();
                var anchor = outputPatterns[0];
                for (var rowIndex = 0; rowIndex < sourceValues.Count; rowIndex++)
                {
                    var row = (TableRow)outputPatterns[rowIndex % outputPatterns.Count].CloneNode(true);
                    for (var columnIndex = 0; columnIndex < request.Columns.Count; columnIndex++)
                    {
                        var cell = CellAtGridColumn(row, request.Columns[columnIndex].TargetGridColumn)!;
                        ReplaceCellTextPreservingTargetStyle(cell, sourceValues[rowIndex][columnIndex]);
                    }
                    outputTable.InsertBefore(row, anchor);
                }
                foreach (var row in outputPatterns) row.Remove();
                outputDocument.MainDocumentPart?.Document?.Save();
                var outputIssues = ValidationIssueCounts(outputDocument);
                var addedIssue = outputIssues.FirstOrDefault(item =>
                    item.Value > baselineValidationIssues.GetValueOrDefault(item.Key));
                if (addedIssue.Key is not null)
                    throw new InvalidOperationException($"output-added-openxml-validation-issues: {addedIssue.Key}");
            }
            File.Move(temporaryPath, outputPath);
            var outputRevision = Observation.CurrentRevision(outputPath);
            var outputRows = ReadBack(
                outputPath,
                outputRevision,
                targetRefs[0],
                targetStartRowIndex,
                request.Columns,
                sourceValues);
            var receipt = new DocxCopyTableRangeReceipt(
                "tiwater.docx-copy-table-range-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                Observation.CurrentRevision(sourcePath),
                Observation.CurrentRevision(targetPath),
                outputRevision,
                request.Source.RowRefs.Count,
                request.Target.RowPatternRefs.Count,
                request.Columns.Count,
                outputRows,
                outputPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
            if (File.Exists(receiptPath)) File.Delete(receiptPath);
            throw;
        }
    }

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path)) throw new InvalidOperationException($"{name}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            throw new InvalidOperationException($"{name}-directory-not-found");
    }

    private static void EnsureContiguousRows(IReadOnlyList<TableRow> rows, string label)
    {
        if (rows.Count == 0 || rows.Any(row => row.Parent is not Table))
            throw new InvalidOperationException($"{label}-rows-not-found");
        var table = (Table)rows[0].Parent!;
        if (rows.Any(row => !ReferenceEquals(row.Parent, table)))
            throw new InvalidOperationException($"{label}-rows-must-share-one-table");
        var tableRows = table.Elements<TableRow>().ToList();
        var indexes = rows.Select(row => tableRows.IndexOf(row)).ToList();
        if (indexes.Any(index => index < 0) || indexes.Distinct().Count() != indexes.Count
            || indexes.Zip(indexes.Skip(1)).Any(pair => pair.Second != pair.First + 1))
            throw new InvalidOperationException($"{label}-rows-must-be-unique-and-contiguous-in-document-order");
    }

    private static bool HasVerticalMerge(IEnumerable<TableRow> rows)
        => rows.SelectMany(row => row.Elements<TableCell>())
            .Any(cell => cell.TableCellProperties?.VerticalMerge is not null);

    private static TableCell? CellAtGridColumn(TableRow row, int gridColumn)
    {
        var current = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            if (gridColumn >= current && gridColumn < current + span) return cell;
            current += span;
        }
        return null;
    }

    private static void EnsureTextOnly(TableCell cell, string label)
    {
        if (cell.Descendants<Drawing>().Any()
            || cell.Descendants<Hyperlink>().Any()
            || cell.Descendants<SimpleField>().Any()
            || cell.Descendants<FieldCode>().Any()
            || cell.Descendants<SdtElement>().Any()
            || cell.Descendants<FootnoteReference>().Any()
            || cell.Descendants<EndnoteReference>().Any())
            throw new InvalidOperationException($"{label}-cell-contains-unsupported-linked-or-structured-content");
    }

    private static void ReplaceCellTextPreservingTargetStyle(TableCell cell, string text)
    {
        var templateParagraph = cell.Elements<Paragraph>().FirstOrDefault();
        var paragraphProperties = templateParagraph?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var runProperties = templateParagraph?.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        cell.RemoveAllChildren<Paragraph>();
        foreach (var line in text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
        {
            var run = new Run();
            if (runProperties is not null) run.RunProperties = (RunProperties)runProperties.CloneNode(true);
            run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
            var paragraph = new Paragraph();
            if (paragraphProperties is not null) paragraph.ParagraphProperties = (ParagraphProperties)paragraphProperties.CloneNode(true);
            paragraph.Append(run);
            cell.Append(paragraph);
        }
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document)
            .GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static IReadOnlyList<DocxCopyTableRangeReadbackRow> ReadBack(
        string output,
        DocxRevision revision,
        ResolvedDocxReference targetTable,
        int startRowIndex,
        IReadOnlyList<DocxCopyTableRangeColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> expectedValues)
    {
        using var document = WordprocessingDocument.Open(output, false);
        var table = Observation.ResolveNativePath(document, targetTable.StoryPart, targetTable.NativePath) as Table
            ?? throw new InvalidOperationException("output-target-table-not-found");
        var rows = table.Elements<TableRow>().ToList();
        if (startRowIndex + expectedValues.Count > rows.Count)
            throw new InvalidOperationException("output-row-range-incomplete");
        var result = new List<DocxCopyTableRangeReadbackRow>(expectedValues.Count);
        for (var rowOffset = 0; rowOffset < expectedValues.Count; rowOffset++)
        {
            var row = rows[startRowIndex + rowOffset];
            var values = columns.Select(mapping =>
            {
                var cell = CellAtGridColumn(row, mapping.TargetGridColumn)
                    ?? throw new InvalidOperationException("output-target-grid-column-not-found");
                return string.Join("\n", cell.Elements<Paragraph>().Select(paragraph => paragraph.InnerText));
            }).ToList();
            if (!values.SequenceEqual(expectedValues[rowOffset], StringComparer.Ordinal))
                throw new InvalidOperationException("output-readback-does-not-match-requested-source-content");
            var nativePath = $"{targetTable.NativePath}/w:tr[{startRowIndex + rowOffset + 1}]";
            result.Add(new DocxCopyTableRangeReadbackRow(
                Observation.MakeReference(revision, "row", targetTable.StoryPart, nativePath),
                nativePath,
                values));
        }
        return result;
    }
}

public sealed record DocxCopyTableRangeRequest(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("source")] DocxCopyTableRangeSource Source,
    [property: JsonPropertyName("target")] DocxCopyTableRangeTarget Target,
    [property: JsonPropertyName("columns")] IReadOnlyList<DocxCopyTableRangeColumn> Columns,
    [property: JsonPropertyName("output")] string Output,
    [property: JsonPropertyName("receiptOutput")] string ReceiptOutput);

public sealed record DocxCopyTableRangeSource(string Input, string Revision, IReadOnlyList<string> RowRefs);
public sealed record DocxCopyTableRangeTarget(string Input, string Revision, string TableRef, IReadOnlyList<string> RowPatternRefs);
public sealed record DocxCopyTableRangeColumn(int SourceGridColumn, int TargetGridColumn);
public sealed record DocxCopyTableRangeReadbackRow(string Ref, string NativePath, IReadOnlyList<string> Values);

public sealed record DocxCopyTableRangeReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    DocxRevision SourceRevision,
    DocxRevision TargetRevision,
    DocxRevision OutputRevision,
    int SourceRowCount,
    int TargetPatternRowCount,
    int ColumnMappingCount,
    IReadOnlyList<DocxCopyTableRangeReadbackRow> OutputRows,
    string Output);
