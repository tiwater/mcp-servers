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
    public const string CopyCommand = "docx_copy_object";
    public const string DeleteCommand = "docx_delete_object";
    private const string MainStory = "/word/document.xml";

    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        var requestJson = File.ReadAllText(args[0]);
        object request;
        object receipt;
        if (command == CopyCommand)
        {
            request = JsonSerializer.Deserialize<CopyObjectRequest>(requestJson, Json.Options)
                ?? throw new InvalidOperationException("copy-object-request-invalid");
            receipt = Copy((CopyObjectRequest)request);
        }
        else if (command == DeleteCommand)
        {
            request = JsonSerializer.Deserialize<DeleteObjectRequest>(requestJson, Json.Options)
                ?? throw new InvalidOperationException("delete-object-request-invalid");
            receipt = Delete((DeleteObjectRequest)request);
        }
        else throw new InvalidOperationException("native-object-command-invalid");
        var output = command == CopyCommand ? ((CopyObjectReceipt)receipt).Output : ((DeleteObjectReceipt)receipt).Output;
        var receiptOutput = command == CopyCommand
            ? ((CopyObjectReceipt)receipt).ReceiptOutput
            : ((DeleteObjectReceipt)receipt).ReceiptOutput;
        var operationCount = command == CopyCommand
            ? ((CopyObjectRequest)request).Changes.Count
            : ((DeleteObjectRequest)request).Changes.Count;
        var appliedCount = command == CopyCommand
            ? ((CopyObjectReceipt)receipt).Changes.Count
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

    public static CopyObjectReceipt Copy(CopyObjectRequest request)
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
            var receipt = new CopyObjectReceipt(
                "tiwater.docx-copy-object-receipt", "tiwater.docx.cli", RuntimeIdentity.Version,
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
                if (group.Key is Table table && group.Count() >= table.Elements<TableRow>().Count())
                    throw new InvalidOperationException("delete-must-not-remove-all-table-rows");
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

    private static readonly HashSet<string> CopyableKinds = ["paragraph", "table", "row", "cell", "run", "text"];
    private static readonly HashSet<string> DeletableKinds = ["paragraph", "table", "row", "cell", "run", "text", "drawing"];

    private static IReadOnlyList<PreparedCopy> PrepareCopies(CopyObjectRequest request, string targetPath)
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
                throw new InvalidOperationException("copy-target-must-be-main-document-object");

            var sourcePath = Path.GetFullPath(change.SourceInput);
            var source = Observation.ResolveAddresses(sourcePath, change.Sources, "changes.sources");
            if (source.Any(item => item.StoryPart != MainStory || !CopyableKinds.Contains(item.Kind)))
                throw new InvalidOperationException("copy-source-kind-not-supported");
            using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
            var clones = source.Select(item => Observation.ResolveNativePath(sourceDocument, item.StoryPart, item.NativePath).CloneNode(true)).ToArray();
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
                        DocxObjectActions.RewriteRelationships(clone, relationshipMap);
                        DocxObjectActions.RemapDrawingIds(outputBody, [clone]);
                        if (before is null) parent.AppendChild(clone); else parent.InsertBefore(clone, before);
                        inserted.Add(clone);
                    }
                }
            }
        }
    }

    private static void EnsureAllowedParent(OpenXmlElement parent, OpenXmlElement child)
    {
        var allowed = parent switch
        {
            Body => child is Paragraph or Table,
            Table => child is TableRow,
            TableRow => child is TableCell,
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
}

public sealed record CopyObjectChange(string SourceInput, IReadOnlyList<DocxObjectAddress> Sources, DocxObjectAddress TargetParent, DocxObjectAddress? Before = null, int? Repeat = null);
public sealed record CopyObjectRequest(string Input, IReadOnlyList<CopyObjectChange> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectChange(IReadOnlyList<DocxObjectAddress> Addresses);
public sealed record DeleteObjectRequest(string Input, IReadOnlyList<DeleteObjectChange> Changes, string Output, string ReceiptOutput);
public sealed record ObjectArtifact(string Path, string Sha256, long Bytes);
public sealed record ObjectReadback(DocxObjectAddress Address, string Kind, string Text);
public sealed record CopyObjectReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<ObjectReadback> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<DocxObjectAddress> DeletedAddresses, string Output, string ReceiptOutput);
