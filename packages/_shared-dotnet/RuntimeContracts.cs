using System.Text.Json;
using System.Text.Json.Serialization;

namespace Tiwater.RuntimeContracts;

public static class RuntimeContractVersions
{
    public const string Capabilities = "1.0.0";
    public const string EvidenceEnvelope = "1.0.0";
    public const string EditReport = "1.0.0";
}

public sealed record PackageIdentity(string Name, string Version);

public sealed record RuntimeIdentity(string Family, string Name, string Version);

public sealed record SchemaIdentity(string Id, string Version);

public sealed record ContractFinding(string Code, string Message, JsonElement? Details = null);

public sealed record FileContentIdentity(
    string Path,
    long SizeBytes,
    string Sha256,
    string ContentId);

public sealed record ArtifactIdentity(
    string ArtifactId,
    long SizeBytes,
    string Sha256,
    string MediaType,
    string Encoding,
    SchemaIdentity Schema);

public sealed record SignatureEvidence(
    string Status,
    string Kind,
    IReadOnlyList<string> Evidence);

public sealed record RuntimeFileEvidence(
    string? FileKind,
    string? MediaType,
    SignatureEvidence Signature);

public sealed record RuntimeEvidenceEnvelope(
    string SchemaVersion,
    string EnvelopeType,
    string Probe,
    string Status,
    string? FailureStage,
    PackageIdentity Package,
    RuntimeIdentity Runtime,
    SchemaIdentity EvidenceSchema,
    FileContentIdentity? Source,
    RuntimeFileEvidence File,
    ArtifactIdentity Artifact,
    JsonElement Payload,
    IReadOnlyList<EvidenceObject> Objects,
    IReadOnlyList<ContractFinding> Warnings,
    IReadOnlyList<ContractFinding> Errors);

public sealed record DiscoveryCommand(
    string Command,
    IReadOnlyList<string> Arguments,
    bool Mutates);

public sealed record IdentifyProbe(
    string Command,
    IReadOnlyList<string> Arguments,
    bool Mutates,
    IReadOnlyList<string> Outcomes);

public sealed record SupportedKind(
    string FileKind,
    IReadOnlyList<string> MediaTypes,
    IReadOnlyList<string> SignatureKinds);

public sealed record RuntimeCommand(
    string Name,
    bool Mutates,
    SchemaIdentity OutputSchema);

public sealed record IdentityPolicy(
    string NativeIds,
    string DerivedIds,
    string Containment);

public sealed record RuntimeCapabilityDescriptor(
    string SchemaVersion,
    string DescriptorType,
    PackageIdentity Package,
    RuntimeIdentity Runtime,
    SchemaIdentity EvidenceSchema,
    DiscoveryCommand DescriptorCommand,
    IdentifyProbe IdentifyProbe,
    IReadOnlyList<SupportedKind> SupportedKinds,
    IReadOnlyList<RuntimeCommand> Commands,
    IdentityPolicy IdentityPolicy);

[JsonPolymorphic(TypeDiscriminatorPropertyName = "kind")]
[JsonDerivedType(typeof(NativeEvidenceIdentity), "native")]
[JsonDerivedType(typeof(DerivedEvidenceIdentity), "derived")]
public abstract record EvidenceIdentity;

public sealed record NativeEvidenceIdentity(
    string Namespace,
    string NativeId) : EvidenceIdentity;

public sealed record DerivedEvidenceIdentity(
    string Derivation,
    IReadOnlyList<string> Inputs) : EvidenceIdentity;

public sealed record EvidenceObject(
    string ObjectId,
    string ObjectType,
    bool Root,
    string? ParentObjectId,
    EvidenceIdentity Identity);

public sealed record TargetReference(string ObjectId);

public sealed record EditOperationResult(
    int Index,
    string Type,
    string Status,
    JsonElement RequestedPayload,
    JsonElement? AppliedPayload,
    IReadOnlyList<TargetReference> Targets,
    IReadOnlyList<ContractFinding> Warnings,
    IReadOnlyList<ContractFinding> Errors)
{
    public static EditOperationResult ForApplied(
        int index,
        string type,
        JsonElement requestedPayload,
        JsonElement appliedPayload,
        IReadOnlyList<TargetReference> targets,
        IReadOnlyList<ContractFinding>? warnings = null) =>
        Create(index, type, "applied", requestedPayload, appliedPayload, targets, warnings, []);

    public static EditOperationResult ForNoop(
        int index,
        string type,
        JsonElement requestedPayload,
        JsonElement appliedPayload,
        IReadOnlyList<TargetReference> targets,
        IReadOnlyList<ContractFinding>? warnings = null) =>
        Create(index, type, "noop", requestedPayload, appliedPayload, targets, warnings, []);

    public static EditOperationResult ForRejected(
        int index,
        string type,
        JsonElement requestedPayload,
        IReadOnlyList<TargetReference> targets,
        IReadOnlyList<ContractFinding> errors,
        IReadOnlyList<ContractFinding>? warnings = null) =>
        Create(index, type, "rejected", requestedPayload, null, targets, warnings, errors);

    public static EditOperationResult ForFailed(
        int index,
        string type,
        JsonElement requestedPayload,
        IReadOnlyList<TargetReference> targets,
        IReadOnlyList<ContractFinding> errors,
        IReadOnlyList<ContractFinding>? warnings = null) =>
        Create(index, type, "failed", requestedPayload, null, targets, warnings, errors);

    private static EditOperationResult Create(
        int index,
        string type,
        string status,
        JsonElement requestedPayload,
        JsonElement? appliedPayload,
        IReadOnlyList<TargetReference> targets,
        IReadOnlyList<ContractFinding>? warnings,
        IReadOnlyList<ContractFinding> errors)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        if (string.IsNullOrWhiteSpace(type)) throw new ArgumentException("Operation type must be non-empty.", nameof(type));
        if ((status is "rejected" or "failed") && errors.Count == 0)
        {
            throw new ArgumentException("Rejected and failed operations require at least one error.", nameof(errors));
        }

        return new EditOperationResult(
            index,
            type,
            status,
            requestedPayload.Clone(),
            appliedPayload?.Clone(),
            targets,
            warnings ?? [],
            errors);
    }
}

public sealed record EditReportSummary(
    int Requested,
    int Applied,
    int Noop,
    int Rejected,
    int Failed)
{
    public static EditReportSummary FromOperations(IReadOnlyList<EditOperationResult> operations)
    {
        for (var index = 0; index < operations.Count; index += 1)
        {
            if (operations[index].Index != index)
            {
                throw new ArgumentException("Edit operation results must preserve zero-based request order.", nameof(operations));
            }
        }

        return new EditReportSummary(
            Requested: operations.Count,
            Applied: operations.Count(operation => operation.Status == "applied"),
            Noop: operations.Count(operation => operation.Status == "noop"),
            Rejected: operations.Count(operation => operation.Status == "rejected"),
            Failed: operations.Count(operation => operation.Status == "failed"));
    }
}

public sealed record RuntimeEditReport(
    string SchemaVersion,
    string ReportType,
    PackageIdentity Package,
    RuntimeIdentity Runtime,
    SchemaIdentity EvidenceSchema,
    FileContentIdentity Source,
    FileContentIdentity Output,
    ArtifactIdentity RequestArtifact,
    IReadOnlyList<EditOperationResult> Operations,
    EditReportSummary Summary,
    IReadOnlyList<ContractFinding> Warnings,
    IReadOnlyList<ContractFinding> Errors);

public static class EditReports
{
    private static readonly SchemaIdentity ReportSchema = new(
        "https://tiwater.dev/contracts/runtime/edit-report.schema.json",
        RuntimeContractVersions.EditReport);
    private static readonly SchemaIdentity RequestSchema = new(
        "tiwater.runtime.edit-request",
        RuntimeContractVersions.EditReport);

    public static RuntimeEditReport Create(
        PackageIdentity package,
        RuntimeIdentity runtime,
        string sourcePath,
        string outputPath,
        JsonElement requestDocument,
        IReadOnlyList<JsonElement> requestOperations,
        IReadOnlyList<EditOperationResult> operations)
    {
        if (requestOperations.Count != operations.Count)
        {
            throw new InvalidOperationException("Edit report operation count must match the authoritative request.");
        }

        for (var index = 0; index < operations.Count; index += 1)
        {
            var result = operations[index];
            var request = requestOperations[index];
            if (result.Index != index)
            {
                throw new InvalidOperationException("Edit report operations must preserve request order.");
            }
            if (request.ValueKind != JsonValueKind.Object
                || !request.TryGetProperty("type", out var type)
                || type.GetString() != result.Type)
            {
                throw new InvalidOperationException($"Edit report operation {index} type does not match the authoritative request.");
            }
            if (!EvidenceEnvelope.CanonicalJsonBytes(request).SequenceEqual(
                    EvidenceEnvelope.CanonicalJsonBytes(result.RequestedPayload)))
            {
                throw new InvalidOperationException($"Edit report operation {index} payload does not match the authoritative request.");
            }
        }

        return new RuntimeEditReport(
            RuntimeContractVersions.EditReport,
            "runtime-edit-report",
            package,
            runtime,
            ReportSchema,
            FileIdentity.IdentifyFile(sourcePath),
            FileIdentity.IdentifyFile(outputPath),
            EvidenceEnvelope.IdentifyCanonicalJson(requestDocument, RequestSchema),
            operations,
            EditReportSummary.FromOperations(operations),
            [],
            []);
    }
}

public static class RuntimeJson
{
    public static JsonSerializerOptions Options { get; } = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };
}
