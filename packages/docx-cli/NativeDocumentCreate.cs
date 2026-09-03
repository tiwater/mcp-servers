using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeDocumentCreate
{
    public const string Command = "docx_create";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CreateDocumentRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("create-document-request-invalid");
        var receipt = Create(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = 1, appliedCount = 1 },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static CreateDocumentReceipt Create(CreateDocumentRequest request)
    {
        var output = Path.GetFullPath(request.Output);
        var receiptOutput = Path.GetFullPath(request.ReceiptOutput);
        if (!string.Equals(Path.GetExtension(output), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("output-must-be-docx");
        if (StringComparer.OrdinalIgnoreCase.Equals(output, receiptOutput))
            throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        RequireNewPath(output, "output");
        RequireNewPath(receiptOutput, "receiptOutput");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        Directory.CreateDirectory(Path.GetDirectoryName(receiptOutput)!);

        var temporary = output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            using (var document = WordprocessingDocument.Create(temporary, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph()));
                main.Document.Save();
                if (NativeMutationSupport.ValidationIssueCounts(document).Count != 0)
                    throw new InvalidOperationException("created-document-openxml-invalid");
            }
            File.Move(temporary, output, false);
            var receipt = new CreateDocumentReceipt(
                "tiwater.docx-create-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                output);
            using (var stream = new FileStream(receiptOutput, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, receipt, Json.CamelCaseOptions);
                stream.WriteByte((byte)'\n');
            }
            return receipt;
        }
        catch
        {
            NativeMutationSupport.Cleanup(temporary, output, receiptOutput);
            throw;
        }
    }

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path) || Directory.Exists(path))
            throw new InvalidOperationException($"{name}-already-exists");
        if (string.IsNullOrWhiteSpace(Path.GetDirectoryName(path)))
            throw new InvalidOperationException($"{name}-directory-invalid");
    }
}

public sealed record CreateDocumentRequest(string Output, string ReceiptOutput);

public sealed record CreateDocumentReceipt(string Schema, string Provider, string ToolVersion, string Output);
