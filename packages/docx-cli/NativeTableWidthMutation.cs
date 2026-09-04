using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTableWidthMutation
{
    public const string Command = "docx_set_table_width";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetTableWidthRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-table-width-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static SetTableWidthReceipt Apply(SetTableWidthRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var duplicate = request.Changes.Select((change, index) => new { change.Table, Index = index })
            .GroupBy(item => item.Table).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
            throw new InvalidOperationException($"table-address-duplicate: changes=[{string.Join(',', duplicate.Select(item => item.Index))}]");
        foreach (var (change, index) in request.Changes.Select((change, index) => (change, index)))
            if (!SupportedType(change.Width.Type) || !uint.TryParse(change.Width.Value, out var value)
                || (change.Width.Type is "dxa" or "pct" ? value == 0 : value != 0))
                throw new InvalidOperationException($"table-width-invalid: changes[{index}].width");

        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveAddresses(paths.Input, request.Changes.Select(change => change.Table).ToArray(), "changes.table");
        for (var index = 0; index < resolved.Count; index++)
            if (resolved[index].Kind != "table")
                throw new InvalidOperationException($"target-must-be-table: changes[{index}].table; kind={resolved[index].Kind}");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                for (var index = 0; index < resolved.Count; index++)
                {
                    var table = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath) as Table
                        ?? throw new InvalidOperationException($"target-must-be-table: changes[{index}].table");
                    var properties = table.GetFirstChild<TableProperties>();
                    if (properties is null) { properties = new TableProperties(); table.PrependChild(properties); }
                    properties.TableWidth = new TableWidth
                    {
                        Type = ParseType(request.Changes[index].Width.Type),
                        Width = request.Changes[index].Width.Value,
                    };
                }
                output.MainDocumentPart!.Document.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            IReadOnlyList<SetTableWidthReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select(item =>
                {
                    var table = (Table)Observation.ResolveNativePath(output, item.StoryPart, item.NativePath);
                    var width = Observation.TableWidthValue(table) ?? throw new InvalidOperationException("output-readback-table-width-missing");
                    return new SetTableWidthReadback(item.Address, width);
                }).ToArray();
            }
            for (var index = 0; index < readback.Count; index++)
                if (readback[index].Width != request.Changes[index].Width)
                    throw new InvalidOperationException("output-readback-table-width-mismatch");
            var receipt = new SetTableWidthReceipt("tiwater.docx-set-table-width-receipt/v1", "tiwater.docx.cli", RuntimeIdentity.Version, readback, paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static bool SupportedType(string type) => type is "auto" or "dxa" or "nil" or "pct";
    private static TableWidthUnitValues ParseType(string type) => type switch
    {
        "dxa" => TableWidthUnitValues.Dxa,
        "pct" => TableWidthUnitValues.Pct,
        "auto" => TableWidthUnitValues.Auto,
        "nil" => TableWidthUnitValues.Nil,
        _ => throw new InvalidOperationException("table-width-type-unsupported"),
    };
}

public sealed record DocxTableWidth(string Type, string Value);
public sealed record SetTableWidthChange(DocxObjectAddress Table, DocxTableWidth Width);
public sealed record SetTableWidthRequest(string Input, IReadOnlyList<SetTableWidthChange> Changes, string Output, string ReceiptOutput);
public sealed record SetTableWidthReadback(DocxObjectAddress Address, DocxTableWidth Width);
public sealed record SetTableWidthReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<SetTableWidthReadback> Changes, string Output);
