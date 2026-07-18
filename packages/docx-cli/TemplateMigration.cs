using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

/// <summary>
/// Exports the immutable object inventories needed before a caller can propose
/// a cross-template migration. It deliberately does not infer business
/// mappings or mutate a document.
/// </summary>
public static class TemplateMigration
{
    private static readonly IReadOnlyDictionary<string, string> EmptyProvenance = new Dictionary<string, string>(StringComparer.Ordinal);
    private static readonly Regex BodyParagraphId = new("^body:paragraph:(?<paragraph>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex HeaderParagraphId = new("^header:(?<header>\\d+):paragraph:(?<paragraph>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex BodyTableCellId = new("^body:table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static int RunAnalyze(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("analyze-template-migration requires <source.docx> <baseline.docx>");
        }

        var json = args.Skip(2).Contains("--json", StringComparer.Ordinal);
        var analysis = Analyze(args[0], args[1]);
        if (json)
        {
            Console.WriteLine(JsonSerializer.Serialize(analysis, Json.Options));
        }
        else
        {
            Console.WriteLine($"Source objects: {analysis.Source.Objects.Count}");
            Console.WriteLine($"Baseline objects: {analysis.Baseline.Objects.Count}");
            Console.WriteLine($"Unresolved findings: {analysis.Findings.Count}");
            Console.WriteLine($"Unsupported object kinds: {string.Join(", ", analysis.UnsupportedObjectKinds)}");
        }
        return 0;
    }

    public static TemplateMigrationAnalysis Analyze(string source, string baseline)
    {
        var sourceInventory = Inventory(source);
        var baselineInventory = Inventory(baseline);
        var baselineById = baselineInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var findings = new List<TemplateMigrationFinding>();

        foreach (var sourceObject in sourceInventory.Objects)
        {
            baselineById.TryGetValue(sourceObject.Id, out var baselineObject);
            var kind = baselineObject is null ? "source-object-unmapped" :
                sourceObject.Kind != baselineObject.Kind ? "object-kind-mismatch" :
                !string.Equals(sourceObject.Text, baselineObject.Text, StringComparison.Ordinal) ? "object-content-differs" :
                !string.Equals(sourceObject.Style, baselineObject.Style, StringComparison.Ordinal) ? "object-style-differs" :
                "object-equivalent";

            if (kind == "object-equivalent")
            {
                continue;
            }

            findings.Add(new TemplateMigrationFinding(
                Id: $"finding:{sourceObject.Id}",
                Kind: kind,
                SourceObjectId: sourceObject.Id,
                BaselineObjectId: baselineObject?.Id,
                Disposition: "requires-declared-mapping",
                Evidence: new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["sourceSha256"] = sourceInventory.Sha256,
                    ["baselineSha256"] = baselineInventory.Sha256,
                    ["sourceKind"] = sourceObject.Kind,
                    ["baselineKind"] = baselineObject?.Kind ?? "missing"
                }));
        }

        return new TemplateMigrationAnalysis(
            Schema: "tiwater.docx.template-migration-analysis/v1",
            Source: sourceInventory,
            Baseline: baselineInventory,
            Findings: findings,
            UnsupportedObjectKinds: ["footnotes", "endnotes", "comments", "content-controls"]);
    }

    public static int RunBuildOperations(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("build-template-migration-operations requires <source.docx> <baseline.docx> <plan.json>");
        }

        var plan = JsonSerializer.Deserialize<TemplateMigrationPlan>(File.ReadAllText(Path.GetFullPath(args[2])), Json.Options)
            ?? throw new InvalidOperationException("template-migration-plan-invalid");
        var result = BuildOperations(args[0], args[1], plan);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    /// <summary>
    /// Turns a caller-declared mapping into the only edit operations permitted
    /// for a migration. This method never selects a source, target, or value.
    /// </summary>
    public static TemplateMigrationOperationBuild BuildOperations(string source, string baseline, TemplateMigrationPlan plan)
    {
        var analysis = Analyze(source, baseline);
        var failures = new List<TemplateMigrationPlanFailure>();
        var operations = new List<DocxEditOperation>();
        var reviewRequired = false;

        if (!string.Equals(plan.Schema, "tiwater.docx.template-migration-plan/v1", StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-schema-invalid", Detail: plan.Schema));
        }
        if (!string.Equals(plan.SourceSha256, analysis.Source.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-source-hash-mismatch", Detail: plan.SourceSha256));
        }
        if (!string.Equals(plan.BaselineSha256, analysis.Baseline.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-hash-mismatch", Detail: plan.BaselineSha256));
        }

        var sourceById = analysis.Source.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var baselineById = analysis.Baseline.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var mappingsBySource = new Dictionary<string, TemplateMigrationMapping>(StringComparer.Ordinal);
        var copyTargets = new HashSet<string>(StringComparer.Ordinal);

        foreach (var mapping in plan.Mappings ?? [])
        {
            if (!sourceById.TryGetValue(mapping.SourceObjectId, out var sourceObject))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-source-object-unknown", mapping.SourceObjectId));
                continue;
            }
            if (!mappingsBySource.TryAdd(mapping.SourceObjectId, mapping))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-source-object-duplicate", mapping.SourceObjectId));
                continue;
            }

            var disposition = mapping.Disposition?.Trim();
            if (string.Equals(disposition, "copy-text", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !baselineById.TryGetValue(mapping.BaselineObjectId, out var baselineObject))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-unknown", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (!string.Equals(sourceObject.Kind, baselineObject.Kind, StringComparison.Ordinal))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-object-kind-mismatch", mapping.SourceObjectId, mapping.BaselineObjectId, $"{sourceObject.Kind}->{baselineObject.Kind}"));
                    continue;
                }
                if (!copyTargets.Add(mapping.BaselineObjectId))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-duplicate", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                var operation = BuildCopyTextOperation(mapping.BaselineObjectId, sourceObject.Text ?? string.Empty);
                if (operation is null)
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-operation-unsupported", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                operations.Add(operation);
            }
            else if (string.Equals(disposition, "review-required", StringComparison.Ordinal) || string.Equals(disposition, "out-of-scope", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.Reason))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-terminal-reason-required", mapping.SourceObjectId, mapping.BaselineObjectId));
                }
                if (string.Equals(disposition, "review-required", StringComparison.Ordinal)) reviewRequired = true;
            }
            else
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-disposition-invalid", mapping.SourceObjectId, mapping.BaselineObjectId, mapping.Disposition));
            }
        }

        foreach (var sourceObject in sourceById.Values.Where(IsContentBearing))
        {
            if (!mappingsBySource.ContainsKey(sourceObject.Id))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-source-content-unmapped", sourceObject.Id));
            }
        }

        var pass = failures.Count == 0 && !reviewRequired;
        var canonicalOperations = pass ? operations : [];
        return new TemplateMigrationOperationBuild(
            Schema: "tiwater.docx.template-migration-operation-build/v1",
            Pass: pass,
            ReviewRequired: reviewRequired,
            SourceSha256: analysis.Source.Sha256,
            BaselineSha256: analysis.Baseline.Sha256,
            OperationsSha256: pass ? HashCanonical(canonicalOperations) : null,
            Operations: canonicalOperations,
            Failures: failures);
    }

    private static TemplateMigrationInventory Inventory(string input)
    {
        var path = Path.GetFullPath(input);
        using var document = WordprocessingDocument.Open(path, false);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var objects = new List<TemplateMigrationObject>();

        AddBodyObjects(objects, body);
        AddHeaderFooterObjects(objects, mainPart.HeaderParts.Select(part => part.Header), "header");
        AddHeaderFooterObjects(objects, mainPart.FooterParts.Select(part => part.Footer), "footer");
        AddDrawingObjects(objects, mainPart.Document, "mainDocument");
        foreach (var (header, index) in mainPart.HeaderParts.Select((part, index) => (part.Header, index))) AddDrawingObjects(objects, header, $"header:{index}");
        foreach (var (footer, index) in mainPart.FooterParts.Select((part, index) => (part.Footer, index))) AddDrawingObjects(objects, footer, $"footer:{index}");

        var styleIndex = 0;
        foreach (var style in mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? [])
        {
            var id = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(id)) continue;
            objects.Add(Object($"style:{styleIndex++}:{id}", "style", "styles", null, style.StyleName?.Val?.Value, id,
                new Dictionary<string, string>(StringComparer.Ordinal) { ["styleType"] = style.Type?.Value.ToString() ?? "unknown" }));
        }

        return new TemplateMigrationInventory(path, HashFile(path), objects);
    }

    private static void AddBodyObjects(List<TemplateMigrationObject> objects, Body body)
    {
        var paragraphIndex = 0;
        var tableIndex = 0;
        var sectionIndex = 0;
        foreach (var child in body.ChildElements)
        {
            if (child is Paragraph paragraph)
            {
                var id = $"body:paragraph:{paragraphIndex++}";
                objects.Add(Object(id, "paragraph", "body", null, Inspector.GetParagraphText(paragraph), ParagraphStyle(paragraph), EmptyProvenance));
                if (paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)
                {
                    objects.Add(Object($"body:section:{sectionIndex++}", "section", "body", id, null, null, EmptyProvenance));
                }
            }
            else if (child is Table table)
            {
                AddTableObjects(objects, table, $"body:table:{tableIndex++}", null, "body");
            }
            else if (child is SectionProperties)
            {
                objects.Add(Object($"body:section:{sectionIndex++}", "section", "body", null, null, null, EmptyProvenance));
            }
        }
    }

    private static void AddTableObjects(List<TemplateMigrationObject> objects, Table table, string tableId, string? parentId, string scope)
    {
        objects.Add(Object(tableId, "table", scope, parentId, null, null, EmptyProvenance));
        foreach (var (row, rowIndex) in table.Elements<TableRow>().Select((row, index) => (row, index)))
        {
            var rowId = $"{tableId}:row:{rowIndex}";
            objects.Add(Object(rowId, "table-row", scope, tableId, null, null, EmptyProvenance));
            foreach (var (cell, cellIndex) in row.Elements<TableCell>().Select((cell, index) => (cell, index)))
            {
                var cellId = $"{rowId}:cell:{cellIndex}";
                objects.Add(Object(cellId, "table-cell", scope, rowId, string.Concat(cell.Elements<Paragraph>().SelectMany(paragraph => paragraph.Descendants<Text>()).Select(text => text.Text)).Trim(), null, EmptyProvenance));
                foreach (var (nestedTable, nestedIndex) in cell.Elements<Table>().Select((nested, index) => (nested, index)))
                {
                    AddTableObjects(objects, nestedTable, $"{cellId}:table:{nestedIndex}", cellId, scope);
                }
            }
        }
    }

    private static void AddHeaderFooterObjects<T>(List<TemplateMigrationObject> objects, IEnumerable<T?> roots, string scope) where T : OpenXmlPartRootElement
    {
        foreach (var (root, rootIndex) in roots.Select((root, index) => (root, index)))
        {
            if (root is null) continue;
            // Table-cell paragraphs are represented by their table-cell object;
            // emitting them here as well would create two mappings for one fact.
            foreach (var (paragraph, paragraphIndex) in root.Elements<Paragraph>().Select((paragraph, index) => (paragraph, index)))
            {
                objects.Add(Object($"{scope}:{rootIndex}:paragraph:{paragraphIndex}", "paragraph", scope, null, Inspector.GetParagraphText(paragraph), ParagraphStyle(paragraph), EmptyProvenance));
            }
            foreach (var (table, tableIndex) in root.Elements<Table>().Select((table, index) => (table, index)))
            {
                AddTableObjects(objects, table, $"{scope}:{rootIndex}:table:{tableIndex}", null, scope);
            }
        }
    }

    private static void AddDrawingObjects(List<TemplateMigrationObject> objects, OpenXmlPartRootElement? root, string scope)
    {
        if (root is null) return;
        foreach (var (_, index) in root.Descendants<Drawing>().Select((drawing, index) => (drawing, index)))
        {
            objects.Add(Object($"{scope}:drawing:{index}", "drawing", scope, null, null, null, EmptyProvenance));
        }
    }

    private static TemplateMigrationObject Object(string id, string kind, string scope, string? parentId, string? text, string? style, IReadOnlyDictionary<string, string> provenance)
        => new(id, kind, scope, parentId, text, style, provenance);

    private static string? ParagraphStyle(Paragraph paragraph) => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    private static bool IsContentBearing(TemplateMigrationObject item)
        => (item.Kind == "paragraph" || item.Kind == "table-cell") && !string.IsNullOrWhiteSpace(item.Text);

    private static DocxEditOperation? BuildCopyTextOperation(string baselineObjectId, string text)
    {
        var bodyParagraph = BodyParagraphId.Match(baselineObjectId);
        if (bodyParagraph.Success)
        {
            return new DocxEditOperation("replaceParagraphText", ParagraphIndex: int.Parse(bodyParagraph.Groups["paragraph"].Value), Text: text);
        }
        var headerParagraph = HeaderParagraphId.Match(baselineObjectId);
        if (headerParagraph.Success)
        {
            return new DocxEditOperation("replaceHeaderParagraphText", HeaderIndex: int.Parse(headerParagraph.Groups["header"].Value), ParagraphIndex: int.Parse(headerParagraph.Groups["paragraph"].Value), Text: text);
        }
        var bodyTableCell = BodyTableCellId.Match(baselineObjectId);
        if (bodyTableCell.Success)
        {
            return new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: int.Parse(bodyTableCell.Groups["table"].Value),
                RowIndex: int.Parse(bodyTableCell.Groups["row"].Value),
                CellIndex: int.Parse(bodyTableCell.Groups["cell"].Value),
                Text: text);
        }
        return null;
    }

    private static string HashCanonical(IReadOnlyList<DocxEditOperation> operations)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(operations, Json.Options))));
    }

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }
}
