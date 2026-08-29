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
            : ((DeleteObjectReceipt)receipt).DeletedRefs.Count;
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
        ValidatePaths(request.TargetDocument.Input, request.Output, request.ReceiptOutput);
        var targetPath = Path.GetFullPath(request.TargetDocument.Input);
        var outputPath = Path.GetFullPath(request.Output);
        var receiptPath = Path.GetFullPath(request.ReceiptOutput);
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
            File.Move(temporaryPath, outputPath);
            var revision = Observation.CurrentRevision(outputPath);
            IReadOnlyList<ObjectReadback> readback;
            using (var output = WordprocessingDocument.Open(outputPath, false))
            {
                readback = inserted.Select(element => Readback(element, revision)).ToArray();
            }
            var receipt = new CopyObjectReceipt(
                "tiwater.docx-copy-object-receipt", "tiwater.docx.cli", RuntimeIdentity.Version,
                Observation.CurrentRevision(targetPath), revision, readback, outputPath, receiptPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            Cleanup(temporaryPath, outputPath, receiptPath);
            throw;
        }
    }

    public static DeleteObjectReceipt Delete(DeleteObjectRequest request)
    {
        ValidatePaths(request.TargetDocument.Input, request.Output, request.ReceiptOutput);
        var targetPath = Path.GetFullPath(request.TargetDocument.Input);
        var outputPath = Path.GetFullPath(request.Output);
        var receiptPath = Path.GetFullPath(request.ReceiptOutput);
        var refs = request.Changes.SelectMany(change => change.Refs).ToArray();
        if (refs.Length == 0 || refs.Distinct(StringComparer.Ordinal).Count() != refs.Length)
            throw new InvalidOperationException("delete-refs-empty-or-duplicate");
        var resolved = Observation.ResolveReferences(targetPath, request.TargetDocument.Revision, refs);
        if (resolved.Any(item => item.StoryPart != MainStory || !DeletableKinds.Contains(item.Kind)))
            throw new InvalidOperationException("delete-ref-kind-not-supported");

        IReadOnlyDictionary<string, int> baseline;
        using (var target = WordprocessingDocument.Open(targetPath, false))
        {
            baseline = ValidationIssueCounts(target);
            var elements = resolved.Select(item => Observation.ResolveNativePath(target, item.StoryPart, item.NativePath)).ToArray();
            if (elements.Any(element => elements.Any(other => !ReferenceEquals(element, other) && element.Ancestors().Contains(other))))
                throw new InvalidOperationException("delete-refs-must-not-overlap");
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
            File.Move(temporaryPath, outputPath);
            var revision = Observation.CurrentRevision(outputPath);
            var receipt = new DeleteObjectReceipt(
                "tiwater.docx-delete-object-receipt", "tiwater.docx.cli", RuntimeIdentity.Version,
                Observation.CurrentRevision(targetPath), revision, resolved.Select(item => item.Reference).ToArray(), outputPath, receiptPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            Cleanup(temporaryPath, outputPath, receiptPath);
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
            var targetRefs = new[] { change.TargetParentRef }.Concat(change.BeforeRef is null ? [] : [change.BeforeRef]).ToArray();
            var target = Observation.ResolveReferences(targetPath, request.TargetDocument.Revision, targetRefs);
            var parent = target[0];
            var before = target.Count == 2 ? target[1] : null;
            if (parent.StoryPart != MainStory || before is not null && before.StoryPart != MainStory)
                throw new InvalidOperationException("copy-target-must-be-main-document-object");

            var sourcePath = Path.GetFullPath(change.SourceDocument.Input);
            var source = Observation.ResolveReferences(sourcePath, change.SourceDocument.Revision, change.SourceRefs);
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
                if (before is not null && !ReferenceEquals(before.Parent, parent)) throw new InvalidOperationException("before-ref-must-be-direct-child-of-target-parent");
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

    private static ObjectReadback Readback(OpenXmlElement element, DocxRevision revision)
    {
        var kind = element switch
        {
            Paragraph => "paragraph", Table => "table", TableRow => "row", TableCell => "cell",
            DocumentFormat.OpenXml.Wordprocessing.Run => "run", Text => "text", _ => element.LocalName,
        };
        var nativePath = Observation.NativePathFor(element);
        return new ObjectReadback(
            Observation.MakeReference(revision, kind, MainStory, nativePath), nativePath, kind, element.InnerText,
            Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(element.OuterXml))).ToLowerInvariant());
    }

    private static void ValidatePaths(string input, string output, string receiptOutput)
    {
        var inputPath = Path.GetFullPath(input);
        var outputPath = Path.GetFullPath(output);
        var receiptPath = Path.GetFullPath(receiptOutput);
        RequireNewPath(outputPath, "output");
        RequireNewPath(receiptPath, "receiptOutput");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, receiptPath)) throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        if (StringComparer.OrdinalIgnoreCase.Equals(inputPath, outputPath)) throw new InvalidOperationException("output-must-not-overwrite-input");
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document).GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static void RejectAddedValidationIssues(WordprocessingDocument document, IReadOnlyDictionary<string, int> baseline)
    {
        var added = ValidationIssueCounts(document).FirstOrDefault(item => item.Value > baseline.GetValueOrDefault(item.Key));
        if (added.Key is not null) throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
    }

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path) || Directory.Exists(path)) throw new InvalidOperationException($"{name}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) throw new InvalidOperationException($"{name}-directory-not-found");
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

    private sealed record PreparedCopy(string SourcePath, ResolvedDocxReference Parent, ResolvedDocxReference? Before, int Repeat, IReadOnlyList<OpenXmlElement> Clones);
}

public sealed record ObjectDocument(string Input, string Revision);
public sealed record CopyObjectChange(ObjectDocument SourceDocument, IReadOnlyList<string> SourceRefs, string TargetParentRef, string? BeforeRef = null, int? Repeat = null);
public sealed record CopyObjectRequest(ObjectDocument TargetDocument, IReadOnlyList<CopyObjectChange> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectChange(IReadOnlyList<string> Refs);
public sealed record DeleteObjectRequest(ObjectDocument TargetDocument, IReadOnlyList<DeleteObjectChange> Changes, string Output, string ReceiptOutput);
public sealed record ObjectArtifact(string Path, string Sha256, long Bytes);
public sealed record ObjectReadback(string Ref, string NativePath, string Kind, string Text, string ContentSha256);
public sealed record CopyObjectReceipt(string Schema, string Provider, string ToolVersion, DocxRevision TargetRevision, DocxRevision OutputRevision, IReadOnlyList<ObjectReadback> Changes, string Output, string ReceiptOutput);
public sealed record DeleteObjectReceipt(string Schema, string Provider, string ToolVersion, DocxRevision TargetRevision, DocxRevision OutputRevision, IReadOnlyList<string> DeletedRefs, string Output, string ReceiptOutput);
