using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativePolicyMutation
{
    public const string FontCommand = "docx_apply_font_policy";
    public const string TocCommand = "docx_apply_toc_style_policy";

    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        if (command == FontCommand)
        {
            var request = JsonSerializer.Deserialize<ApplyFontPolicyRequest>(File.ReadAllText(args[0]), Json.Options)
                ?? throw new InvalidOperationException("apply-font-policy-request-invalid");
            var receipt = ApplyFontPolicy(request);
            WriteResult(command, request.ReceiptOutput, receipt.Output, receipt.AppliedCount);
            return 0;
        }
        if (command == TocCommand)
        {
            var request = JsonSerializer.Deserialize<ApplyTocStylePolicyRequest>(File.ReadAllText(args[0]), Json.Options)
                ?? throw new InvalidOperationException("apply-toc-style-policy-request-invalid");
            var receipt = ApplyTocStylePolicy(request);
            WriteResult(command, request.ReceiptOutput, receipt.Output, receipt.AppliedCount);
            return 0;
        }
        throw new InvalidOperationException("policy-mutation-command-invalid");
    }

    public static PolicyMutationReceipt ApplyFontPolicy(ApplyFontPolicyRequest request)
    {
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var policyBytes = File.ReadAllBytes(Path.GetFullPath(request.Policy));
        var policy = FontPolicy.ReadPolicy(policyBytes);
        if (!FontPolicy.TryNormalize(policy, out var normalized, out var error))
            throw new InvalidOperationException(error ?? "font-policy-invalid");
        var policySha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(policyBytes)).ToLowerInvariant();
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            var bodyCount = 0;
            var tableCount = 0;
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var body = output.MainDocumentPart?.Document?.Body
                    ?? throw new InvalidOperationException("main-document-body-not-found");
                foreach (var run in body.Descendants<Run>().Where(FontPolicy.HasText))
                {
                    var inTable = run.Ancestors<Table>().Any();
                    FontPolicy.Apply(run, inTable ? normalized.Table : normalized.Body);
                    if (inTable) tableCount++; else bodyCount++;
                }
                output.MainDocumentPart!.Document!.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            var validation = FontPolicy.Validate(paths.Output, normalized, policySha256);
            if (!validation.Pass) throw new InvalidOperationException("output-font-policy-readback-failed");
            return WriteReceipt(
                paths,
                "tiwater.docx-apply-font-policy-receipt/v1",
                policySha256,
                bodyCount,
                tableCount,
                bodyCount + tableCount);
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    public static PolicyMutationReceipt ApplyTocStylePolicy(ApplyTocStylePolicyRequest request)
    {
        if (request.IndentCharactersPerLevel < 0)
            throw new InvalidOperationException("indent-characters-per-level-must-be-nonnegative");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var policy = new { request.Italic, request.IndentCharactersPerLevel };
        var policySha256 = NativeMutationSupport.JsonSha256(policy);
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            int matched;
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                matched = TocStylePolicy.Apply(output, request.Italic, request.IndentCharactersPerLevel);
                output.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Save();
                output.MainDocumentPart?.Document?.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            var validation = TocStylePolicy.Validate(paths.Output, request.Italic, request.IndentCharactersPerLevel);
            if (!validation.Pass) throw new InvalidOperationException("output-toc-style-policy-readback-failed");
            return WriteReceipt(
                paths,
                "tiwater.docx-apply-toc-style-policy-receipt/v1",
                policySha256,
                0,
                0,
                matched);
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static PolicyMutationReceipt WriteReceipt(
        NativeMutationSupport.PathsResult paths,
        string schema,
        string policySha256,
        int bodyRunCount,
        int tableRunCount,
        int appliedCount)
    {
        var receipt = new PolicyMutationReceipt(
            schema,
            "tiwater.docx.cli",
            RuntimeIdentity.Version,
            policySha256,
            bodyRunCount,
            tableRunCount,
            appliedCount,
            paths.Output);
        File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
        return receipt;
    }

    private static void WriteResult(string tool, string receipt, string output, int appliedCount)
        => Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool,
            receipt = NativeMutationSupport.Describe(receipt),
            output = NativeMutationSupport.Describe(output),
            summary = new { pass = true, operationCount = 1, appliedCount },
        }, Json.CamelCaseOptions));
}

public sealed record ApplyFontPolicyRequest(string Input, string Policy, string Output, string ReceiptOutput);
public sealed record ApplyTocStylePolicyRequest(string Input, bool Italic, int IndentCharactersPerLevel, string Output, string ReceiptOutput);
public sealed record PolicyMutationReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    string PolicySha256,
    int BodyRunCount,
    int TableRunCount,
    int AppliedCount,
    string Output);
