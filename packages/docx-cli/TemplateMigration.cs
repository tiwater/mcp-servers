using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
    private static readonly Regex FooterParagraphId = new("^footer:(?<footer>\\d+):paragraph:(?<paragraph>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex BodyTableCellId = new("^body:table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex HeaderTableCellId = new("^header:(?<header>\\d+):table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex FooterTableCellId = new("^footer:(?<footer>\\d+):table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

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

    public static int RunDeriveExactTextPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("derive-template-migration-exact-text-plan requires <source.docx> <baseline.docx>");
        }
        var result = DeriveExactTextPlan(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    /// <summary>
    /// A conservative generic mapping strategy for arbitrary layouts. It maps
    /// only a content-bearing source object whose normalized text occurs once
    /// in its source kind and once in the same baseline kind. Everything else
    /// is review-required; it never resolves a duplicate by position.
    /// </summary>
    public static TemplateMigrationMappingDerivation DeriveExactTextPlan(string source, string baseline)
    {
        var analysis = Analyze(source, baseline);
        var sourceContent = analysis.Source.Objects.Where(IsContentBearing).OrderBy(item => item.Id, StringComparer.Ordinal).ToList();
        var baselineContent = analysis.Baseline.Objects.Where(IsContentBearing).ToList();
        var sourceCounts = sourceContent.GroupBy(MappingKey).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var baselineByKey = baselineContent.GroupBy(MappingKey).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var mappings = new List<TemplateMigrationMapping>();
        var unresolved = new List<TemplateMigrationPlanFailure>();

        foreach (var sourceObject in sourceContent)
        {
            var key = MappingKey(sourceObject);
            baselineByKey.TryGetValue(key, out var candidates);
            var sourceCount = sourceCounts[key];
            if (sourceCount == 1 && candidates is { Count: 1 })
            {
                mappings.Add(new TemplateMigrationMapping(sourceObject.Id, candidates[0].Id, "copy-text"));
                continue;
            }

            var reason = candidates is null || candidates.Count == 0
                ? "template-migration-exact-text-target-missing"
                : "template-migration-exact-text-ambiguous";
            mappings.Add(new TemplateMigrationMapping(sourceObject.Id, null, "review-required", reason));
            unresolved.Add(new TemplateMigrationPlanFailure(reason, sourceObject.Id, Detail: $"sourceMatches={sourceCount};baselineMatches={candidates?.Count ?? 0}"));
        }

        foreach (var sourceObject in analysis.Source.Objects.Where(RequiresTerminalMigrationDisposition).OrderBy(item => item.Id, StringComparer.Ordinal))
        {
            mappings.Add(new TemplateMigrationMapping(sourceObject.Id, null, "review-required", "template-migration-automatic-strategy-unsupported"));
            unresolved.Add(new TemplateMigrationPlanFailure("template-migration-automatic-strategy-unsupported", sourceObject.Id, Detail: sourceObject.Kind));
        }

        var plan = new TemplateMigrationPlan(
            Schema: "tiwater.docx.template-migration-plan/v1",
            SourceSha256: analysis.Source.Sha256,
            BaselineSha256: analysis.Baseline.Sha256,
            Mappings: mappings);
        return new TemplateMigrationMappingDerivation(
            Schema: "tiwater.docx.template-migration-exact-text-plan/v1",
            Pass: unresolved.Count == 0,
            Plan: plan,
            Unresolved: unresolved);
    }

    public static int RunDeriveAnchorGapPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("derive-template-migration-anchor-gap-plan requires <source.docx> <baseline.docx>");
        }
        var result = DeriveAnchorGapPlan(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    /// <summary>
    /// Derives review-required semantic candidates only when two consecutive
    /// exact-text anchors enclose equally sized paragraph gaps in the same
    /// document scope. It never turns structural adjacency into an operation.
    /// </summary>
    public static TemplateMigrationMappingDerivation DeriveAnchorGapPlan(string source, string baseline)
    {
        var analysis = Analyze(source, baseline);
        var exact = DeriveExactTextPlan(source, baseline);
        var mappings = exact.Plan.Mappings.ToDictionary(mapping => mapping.SourceObjectId, StringComparer.Ordinal);
        var candidates = new List<(TemplateMigrationObject Source, TemplateMigrationObject Baseline)>();
        foreach (var scope in analysis.Source.Objects.Where(item => item.Kind == "paragraph" && IsContentBearing(item)).Select(item => item.Scope).Distinct(StringComparer.Ordinal))
        {
            var sourceParagraphs = analysis.Source.Objects.Where(item => item.Kind == "paragraph" && item.Scope == scope && IsContentBearing(item)).ToList();
            var baselineParagraphs = analysis.Baseline.Objects.Where(item => item.Kind == "paragraph" && item.Scope == scope && IsContentBearing(item)).ToList();
            candidates.AddRange(FindEqualAnchorGapCandidates(sourceParagraphs, baselineParagraphs, mappings));
        }

        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v1",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            mappings.Values.OrderBy(mapping => mapping.SourceObjectId, StringComparer.Ordinal).ToList());
        var build = BuildOperations(source, baseline, plan);
        var unresolved = new List<TemplateMigrationPlanFailure>(build.Failures);
        foreach (var mapping in plan.Mappings.Where(mapping => string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal)))
        {
            unresolved.Add(new TemplateMigrationPlanFailure(mapping.Reason ?? "template-migration-review-required", mapping.SourceObjectId, mapping.BaselineObjectId));
        }
        foreach (var candidate in candidates)
        {
            unresolved.Add(new TemplateMigrationPlanFailure("template-migration-anchor-gap-candidate-review-required", candidate.Source.Id, candidate.Baseline.Id));
        }
        return new TemplateMigrationMappingDerivation(
            "tiwater.docx.template-migration-anchor-gap-plan/v1",
            build.Pass && unresolved.Count == 0,
            plan,
            unresolved);
    }

    private static IReadOnlyList<(TemplateMigrationObject Source, TemplateMigrationObject Baseline)> FindEqualAnchorGapCandidates(
        IReadOnlyList<TemplateMigrationObject> source,
        IReadOnlyList<TemplateMigrationObject> baseline,
        IDictionary<string, TemplateMigrationMapping> mappings)
    {
        var candidates = new List<(TemplateMigrationObject Source, TemplateMigrationObject Baseline)>();
        var baselineIndexes = baseline.Select((item, index) => (item.Id, index)).ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var rawAnchors = new List<(int SourceIndex, int BaselineIndex)>();
        for (var sourceIndex = 0; sourceIndex < source.Count; sourceIndex += 1)
        {
            if (!mappings.TryGetValue(source[sourceIndex].Id, out var mapping)
                || !string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(mapping.BaselineObjectId)
                || !baselineIndexes.TryGetValue(mapping.BaselineObjectId, out var baselineIndex)) continue;
            rawAnchors.Add((sourceIndex, baselineIndex));
        }
        var anchors = LongestIncreasingAnchorChain(rawAnchors);
        anchors.Insert(0, (-1, -1));
        anchors.Add((source.Count, baseline.Count));
        for (var index = 1; index < anchors.Count; index += 1)
        {
            var previous = anchors[index - 1];
            var next = anchors[index];
            var claimedBaselineIds = mappings.Values
                .Where(mapping => string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(mapping.BaselineObjectId))
                .Select(mapping => mapping.BaselineObjectId!)
                .ToHashSet(StringComparer.Ordinal);
            var sourceGap = source.Skip(previous.SourceIndex + 1).Take(next.SourceIndex - previous.SourceIndex - 1)
                .Where(item => mappings.TryGetValue(item.Id, out var mapping) && string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal))
                .ToList();
            var baselineGap = baseline.Skip(previous.BaselineIndex + 1).Take(next.BaselineIndex - previous.BaselineIndex - 1)
                .Where(item => !claimedBaselineIds.Contains(item.Id))
                .ToList();
            if (sourceGap.Count == 0 || sourceGap.Count != baselineGap.Count) continue;
            for (var gap = 0; gap < sourceGap.Count; gap += 1)
            {
                var sourceObject = sourceGap[gap];
                var baselineObject = baselineGap[gap];
                candidates.Add((sourceObject, baselineObject));
            }
        }
        return candidates;
    }

    private static List<(int SourceIndex, int BaselineIndex)> LongestIncreasingAnchorChain(IReadOnlyList<(int SourceIndex, int BaselineIndex)> anchors)
    {
        var lengths = new int[anchors.Count];
        var previous = Enumerable.Repeat(-1, anchors.Count).ToArray();
        var best = -1;
        for (var current = 0; current < anchors.Count; current += 1)
        {
            lengths[current] = 1;
            for (var candidate = 0; candidate < current; candidate += 1)
            {
                if (anchors[candidate].BaselineIndex >= anchors[current].BaselineIndex || lengths[candidate] + 1 <= lengths[current]) continue;
                lengths[current] = lengths[candidate] + 1;
                previous[current] = candidate;
            }
            if (best < 0 || lengths[current] > lengths[best]) best = current;
        }
        var chain = new List<(int SourceIndex, int BaselineIndex)>();
        for (var index = best; index >= 0; index = previous[index]) chain.Add(anchors[index]);
        chain.Reverse();
        return chain;
    }

    public static int RunResolveSemanticCandidate(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("resolve-template-migration-semantic-candidate requires <source.docx> <baseline.docx> <candidate.json>");
        }
        var candidate = ReadSemanticCandidate(args[2]);
        var result = ResolveSemanticCandidate(args[0], args[1], candidate);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    /// <summary>
    /// Resolves a semantic candidate expressed only as current-document
    /// observable text or media hashes. It never accepts object ids, indexes,
    /// coordinates, or operation payloads from the candidate.
    /// </summary>
    public static TemplateMigrationMappingDerivation ResolveSemanticCandidate(string source, string baseline, TemplateMigrationSemanticCandidate candidate)
    {
        ValidateSemanticCandidate(candidate);
        var analysis = Analyze(source, baseline);
        var automatic = DeriveExactTextPlan(source, baseline);
        var mappings = automatic.Plan.Mappings.ToDictionary(mapping => mapping.SourceObjectId, StringComparer.Ordinal);
        var failures = new List<TemplateMigrationPlanFailure>();

        foreach (var proposal in candidate.Mappings)
        {
            var sourceMatches = ResolveSelector(analysis.Source.Objects, proposal.Source);
            var baselineMatches = ResolveSelector(analysis.Baseline.Objects, proposal.Baseline);
            if (sourceMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceMatches.Count == 0 ? "template-migration-semantic-source-missing" : "template-migration-semantic-source-ambiguous", Detail: proposal.Source.Kind));
                continue;
            }
            if (baselineMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(baselineMatches.Count == 0 ? "template-migration-semantic-baseline-missing" : "template-migration-semantic-baseline-ambiguous", Detail: proposal.Baseline.Kind));
                continue;
            }
            var sourceObject = sourceMatches[0];
            var baselineObject = baselineMatches[0];
            if (!mappings.TryGetValue(sourceObject.Id, out var existing) || !string.Equals(existing.Disposition, "review-required", StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-source-not-pending", sourceObject.Id, baselineObject.Id));
                continue;
            }
            mappings[sourceObject.Id] = new TemplateMigrationMapping(sourceObject.Id, baselineObject.Id, proposal.Disposition, "semantic-candidate-resolved");
        }

        var copiedMediaRelationships = mappings.Values
            .Where(mapping => string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal))
            .Select(mapping => analysis.Source.Objects.Single(item => item.Id == mapping.SourceObjectId))
            .Select(item => item.Provenance.TryGetValue("relationshipId", out var relationshipId) ? relationshipId : null)
            .Where(relationshipId => !string.IsNullOrWhiteSpace(relationshipId))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var drawing in analysis.Source.Objects.Where(item => item.Kind == "drawing"))
        {
            if (drawing.Provenance.TryGetValue("embedRelationshipId", out var relationshipId) && copiedMediaRelationships.Contains(relationshipId))
            {
                mappings.Remove(drawing.Id);
            }
        }

        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v1",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            mappings.Values.OrderBy(mapping => mapping.SourceObjectId, StringComparer.Ordinal).ToList());
        var build = BuildOperations(source, baseline, plan);
        failures.AddRange(build.Failures);
        foreach (var mapping in plan.Mappings.Where(mapping => string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal)))
        {
            failures.Add(new TemplateMigrationPlanFailure(mapping.Reason ?? "template-migration-review-required", mapping.SourceObjectId, mapping.BaselineObjectId));
        }
        return new TemplateMigrationMappingDerivation(
            "tiwater.docx.template-migration-semantic-resolution/v1",
            build.Pass && failures.Count == 0,
            plan,
            failures);
    }

    private static TemplateMigrationSemanticCandidate ReadSemanticCandidate(string file)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(file)));
        ValidateSemanticCandidateJson(document.RootElement);
        return JsonSerializer.Deserialize<TemplateMigrationSemanticCandidate>(document.RootElement.GetRawText(), Json.Options)
            ?? throw new InvalidOperationException("template-migration-semantic-candidate-invalid");
    }

    private static void ValidateSemanticCandidateJson(JsonElement root)
    {
        RequireOnlyFields(root, new HashSet<string>(["schema", "mappings"], StringComparer.Ordinal), "template-migration-semantic-candidate");
        if (!root.TryGetProperty("mappings", out var mappings) || mappings.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-mappings-invalid");
        foreach (var mapping in mappings.EnumerateArray())
        {
            RequireOnlyFields(mapping, new HashSet<string>(["source", "baseline", "disposition"], StringComparer.Ordinal), "template-migration-semantic-candidate-mapping");
            foreach (var side in new[] { "source", "baseline" })
            {
                if (!mapping.TryGetProperty(side, out var selector)) throw new InvalidOperationException($"template-migration-semantic-candidate-{side}-missing");
                RequireOnlyFields(selector, new HashSet<string>(["kind", "scope", "text", "sha256", "parentText", "previousText", "nextText"], StringComparer.Ordinal), $"template-migration-semantic-candidate-{side}");
            }
        }
    }

    private static void RequireOnlyFields(JsonElement element, IReadOnlySet<string> allowed, string label)
    {
        if (element.ValueKind != JsonValueKind.Object) throw new InvalidOperationException($"{label}-object-invalid");
        foreach (var property in element.EnumerateObject()) if (!allowed.Contains(property.Name)) throw new InvalidOperationException($"{label}-unknown-field:{property.Name}");
    }

    private static void ValidateSemanticCandidate(TemplateMigrationSemanticCandidate candidate)
    {
        if (!string.Equals(candidate.Schema, "tiwater.docx.template-migration-semantic-candidate/v1", StringComparison.Ordinal)) throw new InvalidOperationException("template-migration-semantic-candidate-schema-invalid");
        if (candidate.Mappings is null || candidate.Mappings.Count == 0) throw new InvalidOperationException("template-migration-semantic-candidate-mappings-required");
        foreach (var mapping in candidate.Mappings)
        {
            ValidateSemanticSelector(mapping.Source, "source");
            ValidateSemanticSelector(mapping.Baseline, "baseline");
            if (mapping.Disposition is not ("copy-text" or "copy-media")) throw new InvalidOperationException("template-migration-semantic-candidate-disposition-invalid");
        }
    }

    private static void ValidateSemanticSelector(TemplateMigrationSemanticSelector selector, string side)
    {
        if (string.IsNullOrWhiteSpace(selector.Kind)) throw new InvalidOperationException($"template-migration-semantic-{side}-kind-required");
        var text = !string.IsNullOrWhiteSpace(selector.Text);
        var sha = !string.IsNullOrWhiteSpace(selector.Sha256);
        if (text == sha) throw new InvalidOperationException($"template-migration-semantic-{side}-selector-required");
        if (sha && !Regex.IsMatch(selector.Sha256!, "^[A-Fa-f0-9]{64}$", RegexOptions.CultureInvariant)) throw new InvalidOperationException($"template-migration-semantic-{side}-sha256-invalid");
    }

    private static List<TemplateMigrationObject> ResolveSelector(IReadOnlyList<TemplateMigrationObject> objects, TemplateMigrationSemanticSelector selector)
    {
        var normalizedText = selector.Text is null ? null : NormalizeMappingText(selector.Text);
        var normalizedParentText = selector.ParentText is null ? null : NormalizeMappingText(selector.ParentText);
        var normalizedPreviousText = selector.PreviousText is null ? null : NormalizeMappingText(selector.PreviousText);
        var normalizedNextText = selector.NextText is null ? null : NormalizeMappingText(selector.NextText);
        var byId = objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var siblings = objects.Where(item => string.Equals(item.Kind, selector.Kind, StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(selector.Scope) || string.Equals(item.Scope, selector.Scope, StringComparison.Ordinal)))
            .ToList();
        return siblings.Where(item =>
                (normalizedText is null || string.Equals(NormalizeMappingText(item.Text), normalizedText, StringComparison.Ordinal))
                && (selector.Sha256 is null || (item.Provenance.TryGetValue("sha256", out var hash) && string.Equals(hash, selector.Sha256, StringComparison.OrdinalIgnoreCase)))
                && (normalizedParentText is null || (item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent) && string.Equals(NormalizeMappingText(parent.Text), normalizedParentText, StringComparison.Ordinal)))
                && ContextTextMatches(siblings, item, -1, normalizedPreviousText)
                && ContextTextMatches(siblings, item, 1, normalizedNextText))
            .OrderBy(item => item.Id, StringComparer.Ordinal)
            .ToList();
    }

    private static bool ContextTextMatches(IReadOnlyList<TemplateMigrationObject> siblings, TemplateMigrationObject item, int offset, string? expected)
    {
        if (expected is null) return true;
        var index = -1;
        for (var current = 0; current < siblings.Count; current += 1)
        {
            if (!ReferenceEquals(siblings[current], item)) continue;
            index = current;
            break;
        }
        var neighbor = index + offset;
        return neighbor >= 0 && neighbor < siblings.Count
            && string.Equals(NormalizeMappingText(siblings[neighbor].Text), expected, StringComparison.Ordinal);
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
        var mediaCopies = new List<TemplateMigrationMediaCopy>();
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
            else if (string.Equals(disposition, "copy-media", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !baselineById.TryGetValue(mapping.BaselineObjectId, out var baselineObject))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-unknown", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (sourceObject.Kind != "media" || baselineObject.Kind != "media")
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-media-object-required", mapping.SourceObjectId, mapping.BaselineObjectId, $"{sourceObject.Kind}->{baselineObject.Kind}"));
                    continue;
                }
                if (!copyTargets.Add(mapping.BaselineObjectId))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-duplicate", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                mediaCopies.Add(new TemplateMigrationMediaCopy(mapping.SourceObjectId, mapping.BaselineObjectId));
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

        var copiedMediaRelationships = mappingsBySource.Values
            .Where(mapping => string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal))
            .Select(mapping => sourceById[mapping.SourceObjectId])
            .Select(item => item.Provenance.TryGetValue("relationshipId", out var relationshipId) ? relationshipId : null)
            .Where(relationshipId => !string.IsNullOrWhiteSpace(relationshipId))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var sourceObject in sourceById.Values.Where(IsMigrationRequired))
        {
            var drawingCoveredByMedia = sourceObject.Kind == "drawing"
                && sourceObject.Provenance.TryGetValue("embedRelationshipId", out var embeddedRelationshipId)
                && copiedMediaRelationships.Contains(embeddedRelationshipId);
            if (!mappingsBySource.ContainsKey(sourceObject.Id) && !drawingCoveredByMedia)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-source-object-unmapped", sourceObject.Id, Detail: sourceObject.Kind));
            }
        }

        var pass = failures.Count == 0 && !reviewRequired;
        var canonicalOperations = pass ? operations : [];
        var canonicalMediaCopies = pass ? mediaCopies : [];
        return new TemplateMigrationOperationBuild(
            Schema: "tiwater.docx.template-migration-operation-build/v1",
            Pass: pass,
            ReviewRequired: reviewRequired,
            SourceSha256: analysis.Source.Sha256,
            BaselineSha256: analysis.Baseline.Sha256,
            OperationsSha256: pass ? HashCanonical(new { operations = canonicalOperations, mediaCopies = canonicalMediaCopies }) : null,
            Operations: canonicalOperations,
            MediaCopies: canonicalMediaCopies,
            Failures: failures);
    }

    public static int RunApply(string[] args)
    {
        if (args.Length < 4)
        {
            throw new InvalidOperationException("apply-template-migration requires <source.docx> <baseline.docx> <plan.json> <output.docx>");
        }

        var plan = JsonSerializer.Deserialize<TemplateMigrationPlan>(File.ReadAllText(Path.GetFullPath(args[2])), Json.Options)
            ?? throw new InvalidOperationException("template-migration-plan-invalid");
        var result = Apply(args[0], args[1], plan, args[3]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    public static TemplateMigrationApplyResult Apply(string source, string baseline, TemplateMigrationPlan plan, string output)
    {
        var build = BuildOperations(source, baseline, plan);
        if (!build.Pass)
        {
            return new TemplateMigrationApplyResult(
                Schema: "tiwater.docx.template-migration-apply/v1",
                Pass: false,
                Output: null,
                Build: build,
                Edit: null,
                MediaFailures: [],
                Readback: null);
        }

        var outputPath = Path.GetFullPath(output);
        var candidatePath = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory(),
            $".{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.pending");
        var edit = Editor.Apply(Path.GetFullPath(baseline), candidatePath, build.Operations);
        var mediaFailures = ApplyMediaCopies(source, candidatePath, build.MediaCopies);
        var readback = ValidateReadback(source, baseline, candidatePath, plan);
        var pass = edit.AppliedOperations.All(operation => operation.Applied) && mediaFailures.Count == 0 && readback.Pass;
        if (pass)
        {
            File.Move(candidatePath, outputPath, true);
        }
        else if (File.Exists(candidatePath))
        {
            File.Delete(candidatePath);
        }
        return new TemplateMigrationApplyResult(
            Schema: "tiwater.docx.template-migration-apply/v1",
            Pass: pass,
            Output: pass ? outputPath : null,
            Build: build,
            Edit: edit,
            MediaFailures: mediaFailures,
            Readback: readback);
    }

    /// <summary>
    /// Rebuilds both authority inventories and validates the final document;
    /// it does not trust the builder or Editor result as proof of correctness.
    /// </summary>
    public static TemplateMigrationReadback ValidateReadback(string source, string baseline, string output, TemplateMigrationPlan plan)
    {
        var sourceInventory = Inventory(source);
        var baselineInventory = Inventory(baseline);
        var outputInventory = Inventory(output);
        var failures = new List<TemplateMigrationPlanFailure>();

        if (!string.Equals(plan.SourceSha256, sourceInventory.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-source-hash-mismatch"));
        }
        if (!string.Equals(plan.BaselineSha256, baselineInventory.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-hash-mismatch"));
        }

        var sourceById = sourceInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var outputById = outputInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        foreach (var mapping in plan.Mappings ?? [])
        {
            if (!string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal) && !string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal)) continue;
            if (!sourceById.TryGetValue(mapping.SourceObjectId, out var sourceObject) || string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !outputById.TryGetValue(mapping.BaselineObjectId, out var outputObject))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-object-missing", mapping.SourceObjectId, mapping.BaselineObjectId));
                continue;
            }
            if (string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal) && !string.Equals(sourceObject.Text, outputObject.Text, StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-content-mismatch", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
            if (string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal)
                && (!sourceObject.Provenance.TryGetValue("sha256", out var sourceHash)
                    || !outputObject.Provenance.TryGetValue("sha256", out var outputHash)
                    || !string.Equals(sourceHash, outputHash, StringComparison.Ordinal)))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-media-mismatch", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
        }

        var baselineStructure = baselineInventory.Objects
            .Select(StructureFingerprint)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();
        var outputStructure = outputInventory.Objects
            .Select(StructureFingerprint)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();
        if (!baselineStructure.SequenceEqual(outputStructure, StringComparer.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-structure-drift"));
        }

        using (var document = WordprocessingDocument.Open(Path.GetFullPath(output), false))
        {
            var validationErrors = new OpenXmlValidator().Validate(document).Take(10).ToList();
            foreach (var error in validationErrors)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-openxml-invalid", Detail: error.Description));
            }
        }

        return new TemplateMigrationReadback(failures.Count == 0, failures);
    }

    private static IReadOnlyList<TemplateMigrationPlanFailure> ApplyMediaCopies(string source, string output, IReadOnlyList<TemplateMigrationMediaCopy> copies)
    {
        if (copies.Count == 0) return [];
        var failures = new List<TemplateMigrationPlanFailure>();
        var sourceInventory = Inventory(source).Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var outputInventory = Inventory(output).Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        using var sourceDocument = WordprocessingDocument.Open(Path.GetFullPath(source), false);
        using var outputDocument = WordprocessingDocument.Open(Path.GetFullPath(output), true);
        foreach (var copy in copies)
        {
            if (!sourceInventory.TryGetValue(copy.SourceObjectId, out var sourceObject)
                || !outputInventory.TryGetValue(copy.BaselineObjectId, out var outputObject)
                || !TryResolveImagePart(sourceDocument, sourceObject, out var sourceImage)
                || !TryResolveImagePart(outputDocument, outputObject, out var outputImage))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-media-part-unresolved", copy.SourceObjectId, copy.BaselineObjectId));
                continue;
            }
            using var sourceStream = sourceImage.GetStream(FileMode.Open, FileAccess.Read);
            using var outputStream = outputImage.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(outputStream);
        }
        return failures;
    }

    private static bool TryResolveImagePart(WordprocessingDocument document, TemplateMigrationObject item, out ImagePart image)
    {
        image = null!;
        if (item.Kind != "media" || !item.Provenance.TryGetValue("relationshipId", out var relationshipId)) return false;
        var container = ResolvePartContainer(document.MainDocumentPart!, item.Scope);
        if (container?.GetPartById(relationshipId) is not ImagePart part) return false;
        image = part;
        return true;
    }

    private static OpenXmlPartContainer? ResolvePartContainer(MainDocumentPart mainPart, string scope)
    {
        if (scope == "mainDocument") return mainPart;
        var match = Regex.Match(scope, "^(?<kind>header|footer):(?<index>\\d+)$", RegexOptions.CultureInvariant);
        if (!match.Success) return null;
        var index = int.Parse(match.Groups["index"].Value);
        return match.Groups["kind"].Value == "header"
            ? mainPart.HeaderParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ElementAtOrDefault(index)
            : mainPart.FooterParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ElementAtOrDefault(index);
    }

    private static TemplateMigrationInventory Inventory(string input)
    {
        var path = Path.GetFullPath(input);
        using var document = WordprocessingDocument.Open(path, false);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var objects = new List<TemplateMigrationObject>();

        AddBodyObjects(objects, body);
        AddHeaderFooterObjects(objects, mainPart.HeaderParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => part.Header), "header");
        AddHeaderFooterObjects(objects, mainPart.FooterParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => part.Footer), "footer");
        AddDrawingObjects(objects, mainPart.Document, "mainDocument");
        AddRevisionObjects(objects, mainPart.Document, "mainDocument");
        AddMediaObjects(objects, mainPart, "mainDocument");
        foreach (var (headerPart, index) in mainPart.HeaderParts.Select((part, index) => (part, index)))
        {
            AddDrawingObjects(objects, headerPart.Header, $"header:{index}");
            AddRevisionObjects(objects, headerPart.Header, $"header:{index}");
            AddMediaObjects(objects, headerPart, $"header:{index}");
        }
        foreach (var (footerPart, index) in mainPart.FooterParts.Select((part, index) => (part, index)))
        {
            AddDrawingObjects(objects, footerPart.Footer, $"footer:{index}");
            AddRevisionObjects(objects, footerPart.Footer, $"footer:{index}");
            AddMediaObjects(objects, footerPart, $"footer:{index}");
        }

        var styleIndex = 0;
        foreach (var style in mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? [])
        {
            var id = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(id)) continue;
            objects.Add(Object($"style:{styleIndex++}:{id}", "style", "styles", null, style.StyleName?.Val?.Value, id,
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["styleType"] = style.Type?.Value.ToString() ?? "unknown",
                    ["propertiesSha256"] = HashXml(style)
                }));
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
                AddRunObjects(objects, paragraph, id, "body");
                if (paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)
                {
                    var properties = paragraph.ParagraphProperties.GetFirstChild<SectionProperties>()!;
                    objects.Add(Object($"body:section:{sectionIndex++}", "section", "body", id, null, null, SectionProvenance(properties)));
                }
            }
            else if (child is Table table)
            {
                AddTableObjects(objects, table, $"body:table:{tableIndex++}", null, "body");
            }
            else if (child is SectionProperties)
            {
                objects.Add(Object($"body:section:{sectionIndex++}", "section", "body", null, null, null, SectionProvenance((SectionProperties)child)));
            }
        }
    }

    private static void AddTableObjects(List<TemplateMigrationObject> objects, Table table, string tableId, string? parentId, string scope)
    {
        objects.Add(Object(tableId, "table", scope, parentId, null, null, EmptyProvenance));
        foreach (var (row, rowIndex) in table.Elements<TableRow>().Select((row, index) => (row, index)))
        {
            var rowId = $"{tableId}:row:{rowIndex}";
            var rowText = string.Concat(row.Elements<TableCell>().Select(cell => string.Concat(cell.Descendants<Text>().Select(text => text.Text))));
            objects.Add(Object(rowId, "table-row", scope, tableId, rowText, null, EmptyProvenance));
            foreach (var (cell, cellIndex) in row.Elements<TableCell>().Select((cell, index) => (cell, index)))
            {
                var cellId = $"{rowId}:cell:{cellIndex}";
                objects.Add(Object(cellId, "table-cell", scope, rowId, string.Concat(cell.Elements<Paragraph>().SelectMany(paragraph => paragraph.Descendants<Text>()).Select(text => text.Text)).Trim(), null, EmptyProvenance));
                foreach (var (paragraph, paragraphIndex) in cell.Elements<Paragraph>().Select((paragraph, index) => (paragraph, index)))
                {
                    AddRunObjects(objects, paragraph, cellId, scope, paragraphIndex);
                }
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
                var id = $"{scope}:{rootIndex}:paragraph:{paragraphIndex}";
                objects.Add(Object(id, "paragraph", scope, null, Inspector.GetParagraphText(paragraph), ParagraphStyle(paragraph), EmptyProvenance));
                AddRunObjects(objects, paragraph, id, scope);
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
        foreach (var (drawing, index) in root.Descendants<Drawing>().Select((drawing, index) => (drawing, index)))
        {
            var provenance = new Dictionary<string, string>(StringComparer.Ordinal);
            var embeddedRelationshipId = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                .Select(blip => blip.Embed?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (!string.IsNullOrWhiteSpace(embeddedRelationshipId)) provenance["embedRelationshipId"] = embeddedRelationshipId;
            objects.Add(Object($"{scope}:drawing:{index}", "drawing", scope, null, null, null, provenance));
        }
    }

    private static void AddRunObjects(List<TemplateMigrationObject> objects, Paragraph paragraph, string parentId, string scope, int? cellParagraphIndex = null)
    {
        var runPrefix = cellParagraphIndex is null ? parentId : $"{parentId}:paragraph:{cellParagraphIndex.Value}";
        foreach (var (run, index) in paragraph.Descendants<Run>().Select((run, index) => (run, index)))
        {
            var provenance = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["runPropertiesSha256"] = HashXml(run.RunProperties),
                ["paragraphPropertiesSha256"] = HashXml(paragraph.ParagraphProperties)
            };
            var numbering = paragraph.ParagraphProperties?.NumberingProperties;
            if (numbering is not null) provenance["numberingPropertiesSha256"] = HashXml(numbering);
            objects.Add(Object($"{runPrefix}:run:{index}", "run", scope, parentId,
                string.Concat(run.Descendants<Text>().Select(text => text.Text)),
                run.RunProperties?.RunStyle?.Val?.Value, provenance));
        }
    }

    private static void AddRevisionObjects(List<TemplateMigrationObject> objects, OpenXmlPartRootElement? root, string scope)
    {
        if (root is null) return;
        var revisionNames = new HashSet<string>(["ins", "del", "moveFrom", "moveTo"], StringComparer.Ordinal);
        foreach (var (element, index) in root.Descendants().Where(element => revisionNames.Contains(element.LocalName)).Select((element, index) => (element, index)))
        {
            var provenance = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["revisionType"] = element.LocalName,
                ["revisionPropertiesSha256"] = HashXml(element)
            };
            var author = element.GetAttribute("author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            var date = element.GetAttribute("date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            if (!string.IsNullOrWhiteSpace(author)) provenance["author"] = author;
            if (!string.IsNullOrWhiteSpace(date)) provenance["date"] = date;
            objects.Add(Object($"{scope}:revision:{index}", "revision", scope, null, element.InnerText, null, provenance));
        }
    }

    private static void AddMediaObjects(List<TemplateMigrationObject> objects, OpenXmlPartContainer container, string scope)
    {
        foreach (var (image, index) in container.Parts.Select(part => part.OpenXmlPart).OfType<ImagePart>().Select((image, index) => (image, index)))
        {
            using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
            using var sha = SHA256.Create();
            objects.Add(Object($"{scope}:media:{index}", "media", scope, null, null, null,
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["relationshipId"] = container.GetIdOfPart(image),
                    ["contentType"] = image.ContentType,
                    ["sha256"] = Convert.ToHexString(sha.ComputeHash(stream))
                }));
        }
    }

    private static IReadOnlyDictionary<string, string> SectionProvenance(SectionProperties properties)
        => new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["sectionPropertiesSha256"] = HashXml(properties),
            ["pageSizeSha256"] = HashXml(properties.GetFirstChild<PageSize>()),
            ["pageMarginSha256"] = HashXml(properties.GetFirstChild<PageMargin>()),
            ["headerFooterReferencesSha256"] = HashXml(properties.Elements<HeaderReference>().Concat<OpenXmlElement>(properties.Elements<FooterReference>()))
        };

    private static string HashXml(OpenXmlElement? element)
        => element is null ? HashText(string.Empty) : HashText(element.OuterXml);

    private static string HashXml(IEnumerable<OpenXmlElement> elements)
        => HashText(string.Concat(elements.Select(element => element.OuterXml)));

    private static string HashText(string text)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(text)));
    }

    private static TemplateMigrationObject Object(string id, string kind, string scope, string? parentId, string? text, string? style, IReadOnlyDictionary<string, string> provenance)
        => new(id, kind, scope, parentId, text, style, provenance);

    private static string? ParagraphStyle(Paragraph paragraph) => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    private static bool IsContentBearing(TemplateMigrationObject item)
        => (item.Kind == "paragraph" || item.Kind == "table-cell") && !string.IsNullOrWhiteSpace(item.Text);

    private static bool RequiresTerminalMigrationDisposition(TemplateMigrationObject item)
        => item.Kind is "revision" or "drawing" or "media";

    private static bool IsMigrationRequired(TemplateMigrationObject item)
        => IsContentBearing(item) || RequiresTerminalMigrationDisposition(item);

    private static string MappingKey(TemplateMigrationObject item)
        => $"{item.Kind}\u001F{NormalizeMappingText(item.Text)}";

    private static string NormalizeMappingText(string? text)
        => Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim();

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
        var footerParagraph = FooterParagraphId.Match(baselineObjectId);
        if (footerParagraph.Success)
        {
            return new DocxEditOperation("replaceFooterParagraphText", FooterIndex: int.Parse(footerParagraph.Groups["footer"].Value), ParagraphIndex: int.Parse(footerParagraph.Groups["paragraph"].Value), Text: text);
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
        var headerTableCell = HeaderTableCellId.Match(baselineObjectId);
        if (headerTableCell.Success)
        {
            return new DocxEditOperation("replaceHeaderTableCellText", HeaderIndex: int.Parse(headerTableCell.Groups["header"].Value), TableIndex: int.Parse(headerTableCell.Groups["table"].Value), RowIndex: int.Parse(headerTableCell.Groups["row"].Value), CellIndex: int.Parse(headerTableCell.Groups["cell"].Value), Text: text);
        }
        var footerTableCell = FooterTableCellId.Match(baselineObjectId);
        if (footerTableCell.Success)
        {
            return new DocxEditOperation("replaceFooterTableCellText", FooterIndex: int.Parse(footerTableCell.Groups["footer"].Value), TableIndex: int.Parse(footerTableCell.Groups["table"].Value), RowIndex: int.Parse(footerTableCell.Groups["row"].Value), CellIndex: int.Parse(footerTableCell.Groups["cell"].Value), Text: text);
        }
        return null;
    }

    private static string HashCanonical<T>(T value)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value, Json.Options))));
    }

    private static string StructureFingerprint(TemplateMigrationObject item)
        => string.Join("|", item.Id, item.Kind, item.Scope, item.ParentId ?? string.Empty, item.Style ?? string.Empty);

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }
}
