using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeObjectMutation
{
    public const string InsertCommand = "docx_insert_objects";
    public const string DeleteCommand = "docx_delete_object";
    private const string MainStory = "/word/document.xml";
    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        var requestJson = File.ReadAllText(args[0]);
        object request;
        object receipt;
        if (command == InsertCommand)
        {
            request = JsonSerializer.Deserialize<InsertObjectsRequest>(requestJson, Json.Options)
                ?? throw new InvalidOperationException("insert-objects-request-invalid");
            receipt = Insert((InsertObjectsRequest)request);
        }
        else if (command == DeleteCommand)
        {
            request = JsonSerializer.Deserialize<DeleteObjectRequest>(requestJson, Json.Options)
                ?? throw new InvalidOperationException("delete-object-request-invalid");
            receipt = Delete((DeleteObjectRequest)request);
        }
        else throw new InvalidOperationException("native-object-command-invalid");
        var output = command == InsertCommand ? ((InsertObjectsReceipt)receipt).Output : ((DeleteObjectReceipt)receipt).Output;
        var receiptOutput = command == InsertCommand
            ? ((InsertObjectsReceipt)receipt).ReceiptOutput
            : ((DeleteObjectReceipt)receipt).ReceiptOutput;
        var operationCount = command == InsertCommand
            ? ((InsertObjectsRequest)request).Changes.Count
            : ((DeleteObjectRequest)request).Changes.Count;
        var appliedCount = command == InsertCommand
            ? ((InsertObjectsReceipt)receipt).Changes.Count
            : ((DeleteObjectReceipt)receipt).DeletedAddresses.Count;
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = command,
            receipt = Describe(receiptOutput),
            output = Describe(output),
            summary = new { pass = true, operationCount, appliedCount },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static InsertObjectsReceipt Insert(InsertObjectsRequest request)
    {
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var targetPath = paths.Input;
        var outputPath = paths.Output;
        var receiptPath = paths.Receipt;
        var prepared = PrepareCopies(request, targetPath);
        IReadOnlyDictionary<string, int> baseline;
        using (var target = WordprocessingDocument.Open(targetPath, false))
        {
            baseline = ValidationIssueCounts(target);
            PreflightImports(target, prepared);
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(targetPath, temporaryPath, false);
            var inserted = new List<OpenXmlElement>();
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                ApplyCopies(output, prepared, inserted);
                output.MainDocumentPart?.Document?.Save();
                RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            IReadOnlyList<ObjectReadback> readback;
            using (var output = WordprocessingDocument.Open(outputPath, false))
            {
                readback = inserted.Select(Readback).ToArray();
            }
            var receipt = new InsertObjectsReceipt(
                "tiwater.docx-insert-objects-receipt", "tiwater.docx.cli", RuntimeIdentity.Version,
                readback, outputPath, receiptPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    public static DeleteObjectReceipt Delete(DeleteObjectRequest request)
    {
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var targetPath = paths.Input;
        var outputPath = paths.Output;
        var receiptPath = paths.Receipt;
        var addresses = request.Changes.SelectMany(change => change.Addresses).ToArray();
        if (addresses.Length == 0 || addresses.Distinct().Count() != addresses.Length)
            throw new InvalidOperationException("delete-addresses-empty-or-duplicate");
        var resolved = Observation.ResolveAddresses(targetPath, addresses, "changes.addresses");
        if (resolved.Any(item => item.StoryPart != MainStory || !DeletableKinds.Contains(item.Kind)))
            throw new InvalidOperationException("delete-address-kind-not-supported");

        IReadOnlyDictionary<string, int> baseline;
        using (var target = WordprocessingDocument.Open(targetPath, false))
        {
            baseline = ValidationIssueCounts(target);
            var elements = resolved.Select(item => Observation.ResolveNativePath(target, item.StoryPart, item.NativePath)).ToArray();
            if (elements.Any(element => elements.Any(other => !ReferenceEquals(element, other) && element.Ancestors().Contains(other))))
                throw new InvalidOperationException("delete-addresses-must-not-overlap");
            foreach (var group in elements.OfType<TableRow>().GroupBy(row => row.Parent))
                if (group.Key is Table table)
                {
                    if (group.Count() >= table.Elements<TableRow>().Count())
                        throw new InvalidOperationException("delete-must-not-remove-all-table-rows");
                    RequireClosedVerticalMergeGroups(table, group.ToHashSet());
                }
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(targetPath, temporaryPath, false);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var elements = resolved.Select(item => Observation.ResolveNativePath(output, item.StoryPart, item.NativePath)).ToArray();
                foreach (var element in elements.OrderByDescending(element => element.Ancestors().Count())) element.Remove();
                output.MainDocumentPart?.Document?.Save();
                RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            var receipt = new DeleteObjectReceipt(
                "tiwater.docx-delete-object-receipt", "tiwater.docx.cli", RuntimeIdentity.Version,
                resolved.Select(item => item.Address).ToArray(), outputPath, receiptPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static readonly HashSet<string> CopyableKinds = ["paragraph", "table", "row", "run", "text"];
    private static readonly HashSet<string> DeletableKinds = ["paragraph", "table", "row", "run", "text", "drawing"];

    private static IReadOnlyList<PreparedCopy> PrepareCopies(InsertObjectsRequest request, string targetPath)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var result = new List<PreparedCopy>();
        foreach (var change in request.Changes)
        {
            var repeat = change.Repeat ?? 1;
            var targetAddresses = new[] { change.TargetParent }.Concat(change.Before is null ? [] : [change.Before]).ToArray();
            var target = Observation.ResolveAddresses(targetPath, targetAddresses!, "changes.target");
            var parent = target[0];
            var before = target.Count == 2 ? target[1] : null;
            if (parent.StoryPart != MainStory || before is not null && before.StoryPart != MainStory)
                throw new InvalidOperationException("insert-target-must-be-main-document-object");

            var sourcePath = Path.GetFullPath(change.SourceInput);
            var source = Observation.ResolveAddresses(sourcePath, change.Sources, "changes.sources");
            if (source.Any(item => item.StoryPart != MainStory || !CopyableKinds.Contains(item.Kind)))
                throw new InvalidOperationException("insert-source-kind-not-supported");
            if (parent.Kind == "table")
            {
                if (source.Any(item => item.Kind != "row") || before is not null && before.Kind != "row")
                    throw new InvalidOperationException("table-insert-requires-row-sources-and-row-boundary");
                if (before is not null && !ReferenceEquals(before.Element.Parent, parent.Element))
                    throw new InvalidOperationException("before-must-be-direct-child-of-target-parent");
            }
            using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
            var sourceElements = source.Select(item =>
                Observation.ResolveNativePath(sourceDocument, item.StoryPart, item.NativePath)).ToArray();
            if (parent.Kind == "table") ValidateRowSourceSelection(
                sourceElements.Cast<TableRow>().ToArray(),
                (Table)parent.Element,
                before?.Element as TableRow);
            var clones = sourceElements.Select(item => item.CloneNode(true)).ToArray();
            result.Add(new PreparedCopy(sourcePath, parent, before, repeat, clones));
        }
        return result;
    }

    private static void PreflightImports(WordprocessingDocument targetDocument, IReadOnlyList<PreparedCopy> changes)
    {
        var targetMain = targetDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Clones).ToArray();
            foreach (var id in DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal))
                if (!DocxObjectActions.CanCopyRelationship(sourceMain, id, out var error)) throw new InvalidOperationException(error);
            if (!DocxObjectActions.TryImportStyles(sourceMain, targetMain, roots, false, out var styleError)) throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, targetMain, roots, false, out var numberingError)) throw new InvalidOperationException(numberingError);
        }
    }

    private static void ApplyCopies(WordprocessingDocument output, IReadOnlyList<PreparedCopy> changes, List<OpenXmlElement> inserted)
    {
        var outputMain = output.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        var outputBody = outputMain.Document?.Body ?? throw new InvalidOperationException("target-body-not-found");
        var targets = changes.ToDictionary(
            change => change,
            change => (
                Parent: Observation.ResolveNativePath(output, change.Parent.StoryPart, change.Parent.NativePath),
                Before: change.Before is null ? null : Observation.ResolveNativePath(output, change.Before.StoryPart, change.Before.NativePath)));
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var source = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = source.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Clones).ToArray();
            if (!DocxObjectActions.TryImportStyles(sourceMain, outputMain, roots, true, out var styleError)) throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, outputMain, roots, true, out var numberingError)) throw new InvalidOperationException(numberingError);
            var relationshipMap = DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal)
                .ToDictionary(id => id, id => DocxObjectActions.CopyRelationship(sourceMain, outputMain, id), StringComparer.Ordinal);
            foreach (var change in group)
            {
                var (parent, before) = targets[change];
                if (before is not null && !ReferenceEquals(before.Parent, parent)) throw new InvalidOperationException("before-must-be-direct-child-of-target-parent");
                foreach (var template in change.Clones)
                {
                    EnsureAllowedParent(parent, template);
                    for (var count = 0; count < change.Repeat; count++)
                    {
                        var clone = template.CloneNode(true);
                        RemoveCopiedIdentities(clone);
                        DocxObjectActions.RewriteRelationships(clone, relationshipMap);
                        DocxObjectActions.RemapDrawingIds(outputBody, [clone]);
                        if (before is null) parent.AppendChild(clone); else parent.InsertBefore(clone, before);
                        inserted.Add(clone);
                    }
                }
            }
        }
    }

    private static void RemoveCopiedIdentities(OpenXmlElement root)
    {
        foreach (var item in new[] { root }.Concat(root.Descendants()).ToArray())
        {
            item.RemoveAttribute("paraId", Word2010Namespace);
            item.RemoveAttribute("textId", Word2010Namespace);
            if (item is BookmarkStart or BookmarkEnd) item.Remove();
        }
    }

    private static void EnsureAllowedParent(OpenXmlElement parent, OpenXmlElement child)
    {
        var allowed = parent switch
        {
            Body => child is Paragraph or Table,
            Table => child is TableRow,
            TableCell => child is Paragraph or Table,
            Paragraph => child is DocumentFormat.OpenXml.Wordprocessing.Run,
            DocumentFormat.OpenXml.Wordprocessing.Run => child is Text,
            _ => false,
        };
        if (!allowed) throw new InvalidOperationException($"object-kind-cannot-be-child-of-target-parent: {child.LocalName} -> {parent.LocalName}");
    }

    private static ObjectReadback Readback(OpenXmlElement element)
    {
        var kind = element switch
        {
            Paragraph => "paragraph", Table => "table", TableRow => "row", TableCell => "cell",
            DocumentFormat.OpenXml.Wordprocessing.Run => "run", Text => "text", _ => element.LocalName,
        };
        var nativePath = Observation.NativePathFor(element);
        return new ObjectReadback(
            Observation.Address(MainStory, nativePath), kind, element.InnerText);
    }

    private static void ValidateRowSourceSelection(
        IReadOnlyList<TableRow> selectedRows,
        Table targetTable,
        TableRow? before)
    {
        if (selectedRows.Count == 0) throw new InvalidOperationException("row-source-selection-empty");
        var table = selectedRows[0].Parent as Table ?? throw new InvalidOperationException("row-source-table-not-found");
        if (selectedRows.Any(row => !ReferenceEquals(row.Parent, table)))
            throw new InvalidOperationException("row-sources-must-share-one-table");
        var sourceRows = table.Elements<TableRow>().ToArray();
        var indices = selectedRows.Select(row => Array.IndexOf(sourceRows, row)).ToArray();
        if (indices.Any(index => index < 0) || !indices.SequenceEqual(indices.OrderBy(index => index))
            || indices.Zip(indices.Skip(1)).Any(pair => pair.Second != pair.First + 1))
            throw new InvalidOperationException("row-sources-must-be-one-contiguous-native-range");
        if (TableColumnCount(table) != TableColumnCount(targetTable))
            throw new InvalidOperationException("row-source-table-grid-incompatible-with-target");
        var targetRows = targetTable.Elements<TableRow>().ToArray();
        var boundary = before is null ? targetRows.Length : Array.IndexOf(targetRows, before);
        if (boundary < 0) throw new InvalidOperationException("before-row-not-found-in-target-table");
        var activeAtBoundary = VerticalMergeGroups(targetTable).Where(group =>
        {
            var groupIndices = group.Rows.Select(row => Array.IndexOf(targetRows, row)).ToArray();
            return groupIndices.Min() < boundary && boundary <= groupIndices.Max();
        }).Select(group => group.Key).ToHashSet();
        var firstRowMerges = VerticalMergeCells(selectedRows[0]).ToDictionary(cell => cell.Key, cell => cell.Kind);
        var leadingContinuations = firstRowMerges.Where(item => item.Value == MergedCellValues.Continue)
            .Select(item => item.Key).ToHashSet();
        if (!leadingContinuations.IsSubsetOf(activeAtBoundary))
            throw new InvalidOperationException("row-leading-vertical-merge-requires-compatible-target-boundary");
        foreach (var key in activeAtBoundary)
        {
            if (selectedRows.Any(row => !VerticalMergeCells(row).Any(cell =>
                    cell.Key == key && cell.Kind == MergedCellValues.Continue)))
                throw new InvalidOperationException("row-insert-boundary-requires-compatible-vertical-merge");
        }
    }

    private static IReadOnlyList<((int Start, int Span) Key, MergedCellValues Kind)> VerticalMergeCells(TableRow row)
    {
        var result = new List<((int Start, int Span) Key, MergedCellValues Kind)>();
        var cursor = RowOffset(row.TableRowProperties, "gridBefore");
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            var merge = cell.TableCellProperties?.VerticalMerge;
            if (merge is not null)
                result.Add(((cursor, span), merge.Val?.Value ?? MergedCellValues.Continue));
            cursor += span;
        }
        return result;
    }

    private static void RequireClosedVerticalMergeGroups(Table table, IReadOnlySet<TableRow> selected)
    {
        foreach (var group in VerticalMergeGroups(table))
        {
            var selectedCount = group.Rows.Count(selected.Contains);
            if (selectedCount > 0 && selectedCount != group.Rows.Count)
                throw new InvalidOperationException("row-selection-splits-vertical-merge");
        }
    }

    private static IReadOnlyList<VerticalMergeGroup> VerticalMergeGroups(Table table)
    {
        var groups = new List<VerticalMergeGroup>();
        var active = new Dictionary<(int Start, int Span), List<TableRow>>();
        foreach (var row in table.Elements<TableRow>())
        {
            var continued = new HashSet<(int Start, int Span)>();
            var cursor = RowOffset(row.TableRowProperties, "gridBefore");
            foreach (var cell in row.Elements<TableCell>())
            {
                var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                var key = (cursor, span);
                var merge = cell.TableCellProperties?.VerticalMerge;
                if (merge?.Val?.Value == MergedCellValues.Restart)
                {
                    Close(key);
                    active[key] = [row];
                    continued.Add(key);
                }
                else if (merge is not null)
                {
                    if (!active.TryGetValue(key, out var group))
                        throw new InvalidOperationException("vertical-merge-continuation-without-restart");
                    group.Add(row);
                    continued.Add(key);
                }
                cursor += span;
            }
            foreach (var key in active.Keys.Where(key => !continued.Contains(key)).ToArray()) Close(key);
        }
        foreach (var key in active.Keys.ToArray()) Close(key);
        return groups;

        void Close((int Start, int Span) key)
        {
            if (!active.Remove(key, out var group)) return;
            groups.Add(new VerticalMergeGroup(key, group));
        }
    }

    private static int RowOffset(TableRowProperties? properties, string localName)
    {
        var value = properties?.ChildElements.FirstOrDefault(child => child.LocalName == localName)
            ?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return int.TryParse(value, out var result) ? result : 0;
    }

    private static int TableColumnCount(Table table)
        => Math.Max(
            table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0,
            table.Elements<TableRow>().Select(row =>
                    RowOffset(row.TableRowProperties, "gridBefore")
                    + row.Elements<TableCell>().Sum(cell =>
                        Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1))
                    + RowOffset(row.TableRowProperties, "gridAfter"))
                .DefaultIfEmpty(0).Max());

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document).GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static void RejectAddedValidationIssues(WordprocessingDocument document, IReadOnlyDictionary<string, int> baseline)
    {
        var added = ValidationIssueCounts(document).FirstOrDefault(item => item.Value > baseline.GetValueOrDefault(item.Key));
        if (added.Key is not null) throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
    }

    private static void Cleanup(params string[] paths)
    {
        foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
    }

    private static ObjectArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new ObjectArtifact(Path.GetFullPath(path), Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
    }

    private sealed record PreparedCopy(string SourcePath, ResolvedDocxAddress Parent, ResolvedDocxAddress? Before, int Repeat, IReadOnlyList<OpenXmlElement> Clones);
    private sealed record VerticalMergeGroup((int Start, int Span) Key, IReadOnlyList<TableRow> Rows);
}

public sealed record InsertObjectsChange(string SourceInput, IReadOnlyList<DocxObjectAddress> Sources, DocxObjectAddress TargetParent, DocxObjectAddress? Before = null, int? Repeat = null);
public sealed record InsertObjectsRequest(string Input, IReadOnlyList<InsertObjectsChange> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectChange(IReadOnlyList<DocxObjectAddress> Addresses);
public sealed record DeleteObjectRequest(string Input, IReadOnlyList<DeleteObjectChange> Changes, string Output, string ReceiptOutput);
public sealed record ObjectArtifact(string Path, string Sha256, long Bytes);
public sealed record ObjectReadback(DocxObjectAddress Address, string Kind, string Text);
public sealed record InsertObjectsReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<ObjectReadback> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<DocxObjectAddress> DeletedAddresses, string Output, string ReceiptOutput);
