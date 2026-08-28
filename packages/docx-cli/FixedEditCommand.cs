using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Docx;

public static class FixedEditCommand
{
    private sealed record Definition(string OperationType, bool Batch, string[] SourceFields);
    private sealed record Artifact(string Path, string Sha256, long Bytes);

    private static readonly IReadOnlyDictionary<string, Definition> Definitions =
        new Dictionary<string, Definition>(StringComparer.Ordinal)
    {
        ["docx_set_anchored_text"] = new("replaceAnchoredText", true, []),
        ["docx_set_paragraph_text"] = new("replaceParagraphText", true, []),
        ["docx_set_paragraph_run_text"] = new("replaceParagraphRunText", true, []),
        ["docx_replace_body_text"] = new("replaceBodyText", true, []),
        ["docx_delete_body_paragraph"] = new("deleteBodyParagraph", true, []),
        ["docx_delete_body_drawing_before_paragraph"] = new("deleteBodyDrawingBeforeParagraph", true, []),
        ["docx_insert_body_range"] = new("insertBodyRange", true, ["source"]),
        ["docx_replace_drawing_image"] = new("replaceDrawingImage", true, ["image"]),
        ["docx_insert_body_image"] = new("insertBodyImage", true, ["image"]),
        ["docx_delete_body_range"] = new("deleteBodyRange", true, []),
        ["docx_start_section"] = new("startSectionBeforeParagraph", true, []),
        ["docx_set_header_paragraph_text"] = new("replaceHeaderParagraphText", true, []),
        ["docx_set_header_run_text"] = new("replaceHeaderParagraphRunText", true, []),
        ["docx_replace_header_text"] = new("replaceHeaderText", true, []),
        ["docx_set_footer_paragraph_text"] = new("replaceFooterParagraphText", true, []),
        ["docx_set_footer_run_text"] = new("replaceFooterParagraphRunText", true, []),
        ["docx_set_table_cell_text"] = new("replaceTableCellText", true, []),
        ["docx_set_table_cell_run_text"] = new("replaceTableCellRunText", true, []),
        ["docx_set_header_table_cell_text"] = new("replaceHeaderTableCellText", true, []),
        ["docx_set_header_table_cell_run_text"] = new("replaceHeaderTableCellRunText", true, []),
        ["docx_set_footer_table_cell_text"] = new("replaceFooterTableCellText", true, []),
        ["docx_set_footer_table_cell_run_text"] = new("replaceFooterTableCellRunText", true, []),
        ["docx_set_table_cell_rich_text"] = new("replaceTableCellRichText", true, []),
        ["docx_insert_table_rows"] = new("insertTableRows", true, []),
        ["docx_delete_table_rows"] = new("deleteTableRows", true, []),
        ["docx_replace_table_rows"] = new("replaceTableRows", true, []),
        ["docx_insert_table_columns"] = new("insertTableColumns", true, []),
        ["docx_set_table_width"] = new("setTableWidth", true, []),
        ["docx_set_table_cell_alignment"] = new("setTableCellAlignment", true, []),
        ["docx_set_table_cell_no_wrap"] = new("setTableCellNoWrap", true, []),
        ["docx_set_table_cell_font_size"] = new("setTableCellFontSize", true, []),
        ["docx_apply_font_policy"] = new("applyDocumentFontPolicy", true, []),
        ["docx_set_table_row_height"] = new("setTableRowHeight", true, []),
        ["docx_set_table_row_cant_split"] = new("setTableRowCantSplit", true, []),
        ["docx_set_table_row_repeat_as_header"] = new("setTableRowRepeatAsHeader", true, []),
        ["docx_set_table_row_keep_next"] = new("setTableRowKeepNext", true, []),
        ["docx_set_body_paragraph_keep_next"] = new("setBodyParagraphKeepNext", true, []),
        ["docx_set_body_paragraph_keep_lines"] = new("setBodyParagraphKeepLines", true, []),
        ["docx_apply_toc_style_policy"] = new("applyTocStylePolicy", true, []),
        ["docx_set_header_paragraph_font_size"] = new("setHeaderParagraphFontSize", true, []),
        ["docx_collapse_trailing_empty_section"] = new("collapseTrailingEmptySection", false, []),
        ["docx_collapse_trailing_empty_paragraphs"] = new("collapseTrailingEmptyBodyParagraphs", false, []),
        ["docx_merge_table_cells"] = new("mergeTableCells", true, []),
        ["docx_unmerge_table_row_cells"] = new("unmergeTableRowHorizontalCells", true, []),
        ["docx_unmerge_table_column_cells"] = new("unmergeTableColumnVerticalCells", true, []),
        ["docx_delete_comments"] = new("deleteComments", true, []),
        ["docx_mark_fields_dirty"] = new("markFieldsDirty", false, []),
        ["docx_sanitize_fields"] = new("sanitizeFields", false, []),
        ["docx_freeze_fields"] = new("freezeFields", false, []),
    };

    public static IReadOnlyCollection<string> Commands => Definitions.Keys.ToArray();
    public static bool IsCommand(string command) => Definitions.ContainsKey(command);

    public static int Run(string command, string[] args)
    {
        if (!Definitions.TryGetValue(command, out var definition))
            throw new InvalidOperationException($"Unknown fixed DOCX command: {command}");
        if (args.Length != 1)
            throw new InvalidOperationException($"{command} requires <request.json>");

        var root = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
            ?? throw new InvalidOperationException("fixed-edit-request-invalid");
        var input = RequirePath(root, "input");
        var output = RequirePath(root, "output");
        var receiptOutput = RequirePath(root, "receiptOutput");
        RequireNewPath(output, "output");
        RequireNewPath(receiptOutput, "receiptOutput");
        if (output == receiptOutput) throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");

        var changes = root["changes"] as JsonArray;
        if (definition.Batch && (changes is null || changes.Count == 0))
            throw new InvalidOperationException("changes-must-contain-at-least-one-item");
        if (!definition.Batch && changes is not null)
            throw new InvalidOperationException("changes-not-accepted-by-document-action");

        var operations = BuildOperations(definition, changes);
        var inputArtifact = Describe(input);
        var sourcePaths = SourcePaths(definition, changes);
        var sourceArtifacts = sourcePaths.Select(Describe).ToArray();
        try
        {
            var editResult = Editor.Apply(input, output, operations);
            var observedSources = sourcePaths.Select(Describe).ToArray();
            var sourceBindingStable = sourceArtifacts.SequenceEqual(observedSources);
            var pass = sourceBindingStable
                && editResult.AppliedOperations.Count == operations.Count
                && editResult.AppliedOperations.All(operation => operation.Applied);
            var outputArtifact = pass ? Describe(output) : null;
            if (!pass && File.Exists(output)) File.Delete(output);
            var applied = editResult.AppliedOperations.Select((operation, index) => new
            {
                index,
                applied = operation.Applied,
                detail = operation.Detail,
            }).ToArray();
            var receiptPayload = new
            {
                schema = "tiwater.office.fixed-edit-receipt/v2",
                tool = command,
                pass,
                input = inputArtifact,
                sources = sourceArtifacts,
                sourceBindingStable,
                acceptedCall = root,
                output = outputArtifact,
                operationCount = operations.Count,
                appliedOperations = applied,
            };
            var receipt = WriteJsonArtifact(receiptOutput, receiptPayload);
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                tool = command,
                receipt,
                output = outputArtifact,
                summary = new
                {
                    pass,
                    operationCount = operations.Count,
                    appliedCount = applied.Count(operation => operation.applied),
                },
            }, Json.CamelCaseOptions));
            return pass ? 0 : 1;
        }
        catch
        {
            if (File.Exists(output)) File.Delete(output);
            throw;
        }
    }

    private static IReadOnlyList<DocxEditOperation> BuildOperations(Definition definition, JsonArray? changes)
    {
        var values = definition.Batch ? changes! : new JsonArray(new JsonObject());
        return values.Select(change =>
        {
            var operation = change?.DeepClone() as JsonObject
                ?? throw new InvalidOperationException("change-must-be-an-object");
            operation["type"] = definition.OperationType;
            return operation.Deserialize<DocxEditOperation>(Json.Options)
                ?? throw new InvalidOperationException("change-does-not-match-fixed-operation");
        }).ToArray();
    }

    private static string[] SourcePaths(Definition definition, JsonArray? changes) =>
        definition.SourceFields.SelectMany(field => (changes ?? []).Select(change =>
            Path.GetFullPath(change?[field]?.GetValue<string>()
                ?? throw new InvalidOperationException($"{field}-is-required"))))
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    private static string RequirePath(JsonObject root, string property)
    {
        var value = root[property]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException($"{property}-is-required");
        return Path.GetFullPath(value);
    }

    private static void RequireNewPath(string value, string property)
    {
        if (File.Exists(value) || Directory.Exists(value))
            throw new InvalidOperationException($"{property}-already-exists");
        Directory.CreateDirectory(Path.GetDirectoryName(value) ?? Directory.GetCurrentDirectory());
    }

    private static Artifact Describe(string file)
    {
        using var stream = File.OpenRead(file);
        return new Artifact(Path.GetFullPath(file), Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
    }

    private static Artifact WriteJsonArtifact<T>(string file, T value)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(value, Json.CamelCaseOptions);
        using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            stream.Write(bytes);
            stream.WriteByte((byte)'\n');
        }
        return Describe(file);
    }
}
