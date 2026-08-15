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
    private static readonly Regex BodyParagraphRunId = new("^body:paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex HeaderParagraphRunId = new("^header:(?<header>\\d+):paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex FooterParagraphRunId = new("^footer:(?<footer>\\d+):paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex BodyTableCellRunId = new("^body:table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+):paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex HeaderTableCellRunId = new("^header:(?<header>\\d+):table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+):paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex FooterTableCellRunId = new("^footer:(?<footer>\\d+):table:(?<table>\\d+):row:(?<row>\\d+):cell:(?<cell>\\d+):paragraph:(?<paragraph>\\d+):run:(?<run>\\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
                Disposition: "requires-semantic-candidate",
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
            UnsupportedObjectKinds: sourceInventory.Objects
                .Where(IsUnsupportedObject)
                .Select(item => item.Kind)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(kind => kind, StringComparer.Ordinal)
                .ToList());
    }

    public static int RunDeriveExactTextPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("derive-template-migration-exact-text-plan requires <source.docx> <baseline.docx>");
        }
        var result = DeriveExactTextPlan(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    /// <summary>
    /// A conservative generic mapping strategy for arbitrary layouts. It maps
    /// content that is unique within both inventories. Repeated table-cell
    /// content is mapped only when the complete normalized cell topology makes
    /// the source and baseline tables a reciprocally unique semantic pair;
    /// caller-supplied positions are never accepted.
    /// </summary>
    public static TemplateMigrationMappingDerivation DeriveExactTextPlan(string source, string baseline)
    {
        var analysis = Analyze(source, baseline);
        return DeriveExactTextPlan(analysis);
    }

    private static TemplateMigrationMappingDerivation DeriveExactTextPlan(TemplateMigrationAnalysis analysis)
    {
        var sourceContent = analysis.Source.Objects.Where(IsContentBearing).OrderBy(item => item.Id, StringComparer.Ordinal).ToList();
        var baselineContent = analysis.Baseline.Objects.Where(IsContentBearing).ToList();
        var sourceCounts = sourceContent.GroupBy(MappingKey).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var baselineByKey = baselineContent.GroupBy(MappingKey).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var reciprocalTableTargets = DeriveReciprocalTableCellTargets(analysis);
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
            if (candidates is { Count: > 0 }
                && reciprocalTableTargets.TryGetValue(sourceObject.Id, out var reciprocalTarget)
                && candidates.Any(candidate => string.Equals(candidate.Id, reciprocalTarget, StringComparison.Ordinal)))
            {
                mappings.Add(new TemplateMigrationMapping(sourceObject.Id, reciprocalTarget, "copy-text"));
                continue;
            }

            var reason = candidates is null || candidates.Count == 0
                ? "template-migration-exact-text-match-missing"
                : "template-migration-exact-text-match-non-unique";
            mappings.Add(new TemplateMigrationMapping(sourceObject.Id, null, "unresolved", reason));
            unresolved.Add(new TemplateMigrationPlanFailure(
                reason,
                sourceObject.Id,
                Detail: $"sourceMatches={sourceCount};baselineMatches={candidates?.Count ?? 0}",
                Source: ObserveForSemanticDecision(sourceObject),
                BaselineOptions: candidates?.Select(ObserveForSemanticDecision).ToList() ?? []));
        }

        var sourceMedia = analysis.Source.Objects.Where(item => item.Kind == "media").OrderBy(item => item.Id, StringComparer.Ordinal).ToList();
        var baselineMedia = analysis.Baseline.Objects.Where(item => item.Kind == "media").ToList();
        var sourceMediaByHash = sourceMedia
            .Where(item => item.Provenance.ContainsKey("sha256"))
            .GroupBy(item => item.Provenance["sha256"], StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var baselineMediaByHash = baselineMedia
            .Where(item => item.Provenance.ContainsKey("sha256"))
            .GroupBy(item => item.Provenance["sha256"], StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        foreach (var sourceObject in sourceMedia)
        {
            sourceObject.Provenance.TryGetValue("sha256", out var contentHash);
            var sourceMatches = contentHash is not null && sourceMediaByHash.TryGetValue(contentHash, out var sourceHashMatches)
                ? sourceHashMatches.Count
                : 0;
            var baselineMatches = contentHash is not null && baselineMediaByHash.TryGetValue(contentHash, out var baselineHashMatches)
                ? baselineHashMatches
                : null;
            if (sourceMatches == 1 && baselineMatches is { Count: 1 })
            {
                mappings.Add(new TemplateMigrationMapping(sourceObject.Id, baselineMatches[0].Id, "copy-media"));
                continue;
            }

            var reason = baselineMatches is null || baselineMatches.Count == 0
                ? "template-migration-media-hash-target-missing"
                : "template-migration-media-hash-ambiguous";
            mappings.Add(new TemplateMigrationMapping(sourceObject.Id, null, "unresolved", reason));
            unresolved.Add(new TemplateMigrationPlanFailure(
                reason,
                sourceObject.Id,
                Detail: $"sourceMatches={sourceMatches};baselineMatches={baselineMatches?.Count ?? 0}",
                Source: ObserveForSemanticDecision(sourceObject),
                BaselineOptions: baselineMatches?.Select(ObserveForSemanticDecision).ToList() ?? []));
        }

        var coveredDrawingRelationships = DeriveCoveredDrawingRelationships(analysis, mappings);

        foreach (var sourceObject in analysis.Source.Objects.Where(RequiresTerminalMigrationDisposition).OrderBy(item => item.Id, StringComparer.Ordinal))
        {
            var handledMediaObject = sourceObject.Kind == "media";
            var coveredDrawing = sourceObject.Kind == "drawing"
                && MediaRelationshipKey(sourceObject, "embedRelationshipId") is { } relationshipKey
                && coveredDrawingRelationships.Contains(relationshipKey);
            if (handledMediaObject || coveredDrawing) continue;
            mappings.Add(new TemplateMigrationMapping(sourceObject.Id, null, "unresolved", "template-migration-automatic-strategy-unsupported"));
            unresolved.Add(new TemplateMigrationPlanFailure(
                "template-migration-automatic-strategy-unsupported",
                sourceObject.Id,
                Detail: sourceObject.Kind,
                Source: ObserveForSemanticDecision(sourceObject)));
        }

        var plan = new TemplateMigrationPlan(
            Schema: "tiwater.docx.template-migration-plan/v1",
            SourceSha256: analysis.Source.Sha256,
            BaselineSha256: analysis.Baseline.Sha256,
            Mappings: mappings);
        return new TemplateMigrationMappingDerivation(
            Schema: "tiwater.docx.template-migration-exact-text-plan/v1",
            Pass: true,
            Plan: plan,
            Unresolved: unresolved,
            UnclaimedBaseline: ObserveUnclaimedBaseline(analysis, plan));
    }

    private static IReadOnlyDictionary<string, string> DeriveReciprocalTableCellTargets(TemplateMigrationAnalysis analysis)
    {
        var sourceTables = BuildSemanticTableTopologies(analysis.Source.Objects);
        var baselineTables = BuildSemanticTableTopologies(analysis.Baseline.Objects);
        var sourceBySignature = sourceTables.GroupBy(table => table.Signature, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var baselineBySignature = baselineTables.GroupBy(table => table.Signature, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var (signature, sourceMatches) in sourceBySignature)
        {
            if (sourceMatches.Count != 1
                || !baselineBySignature.TryGetValue(signature, out var baselineMatches)
                || baselineMatches.Count != 1) continue;
            var sourceTable = sourceMatches[0];
            var baselineTable = baselineMatches[0];
            foreach (var (topology, sourceCell) in sourceTable.Cells)
            {
                if (baselineTable.Cells.TryGetValue(topology, out var baselineCell))
                {
                    targets[sourceCell.Id] = baselineCell.Id;
                }
            }
        }
        return targets;
    }

    private static IReadOnlyList<SemanticTableTopology> BuildSemanticTableTopologies(IReadOnlyList<TemplateMigrationObject> objects)
    {
        var tables = new List<SemanticTableTopology>();
        foreach (var group in objects.Where(item => item.Kind == "table-cell" && item.Topology is not null)
            .GroupBy(item => item.Topology!.ContainerObjectId, StringComparer.Ordinal))
        {
            var cells = group.OrderBy(item => item.Topology!.Row).ThenBy(item => item.Topology!.Column).ToList();
            var positions = cells.Select(item => (item.Topology!.Row, item.Topology!.Column)).ToList();
            if (positions.Distinct().Count() != positions.Count) continue;
            var scopes = cells.Select(item => item.Scope).Distinct(StringComparer.Ordinal).ToList();
            if (scopes.Count != 1) continue;
            var signature = scopes[0] + "\u001E" + string.Join("\u001E", cells.Select(item =>
            {
                var text = NormalizeMappingText(item.Text);
                return $"{item.Topology!.Row}:{item.Topology.Column}:{text.Length}:{text}";
            }));
            tables.Add(new SemanticTableTopology(
                signature,
                cells.ToDictionary(item => (item.Topology!.Row, item.Topology.Column), item => item)));
        }
        return tables;
    }

    private sealed record SemanticTableTopology(
        string Signature,
        IReadOnlyDictionary<(int Row, int Column), TemplateMigrationObject> Cells);

    public static int RunDeriveAnchorGapPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("derive-template-migration-anchor-gap-plan requires <source.docx> <baseline.docx>");
        }
        var result = DeriveAnchorGapPlan(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static int RunFindCandidates(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("list-template-migration-options requires <source.docx> <baseline.docx>");
        }
        var result = FindCandidates(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static int RunListChoices(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("list-template-migration-choices requires <source.docx> <baseline.docx>");
        }
        var result = ListChoices(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.CamelCaseOptions));
        return 0;
    }

    public static TemplateMigrationChoiceCatalog ListChoices(string source, string baseline)
    {
        var discovery = FindCandidates(source, baseline);
        var sources = discovery.RequiredDecisions
            .Select(item => ToChoice(
                ChoiceId("source", discovery.SourceSha256, item.Source.Selector
                    ?? throw new InvalidOperationException("template-migration-choice-source-selector-missing")),
                item.Source,
                item.Count,
                item.RequiredCardinality,
                ["place-content", "keep-template-content", "keep-template-label", "select-template-option", "exclude-source", "review-source"]))
            .ToList();
        EnsureUniqueChoiceIds(sources, "source");

        var targetObservations = discovery.AvailableTargets
            .SelectMany(item => new[] { item }.Concat(item.Context?.SelectableChildren ?? []))
            .ToList();
        var targets = targetObservations
            .GroupBy(item => ChoiceId("target", discovery.BaselineSha256, item.Selector
                ?? throw new InvalidOperationException("template-migration-choice-target-selector-missing")), StringComparer.Ordinal)
            .Select(group => ToChoice(
                group.Key,
                group.First(),
                group.Count(),
                allowedActions: group.First().Kind == "run"
                    ? ["select-template-option"]
                    : group.First().Kind == "media"
                        ? ["place-content"]
                    : group.First().Kind == "table-cell"
                        ? ["place-content", "keep-template-content", "keep-template-label", "template-cleanup"]
                        : ["place-content", "keep-template-content", "keep-template-label"]))
            .ToList();

        return new TemplateMigrationChoiceCatalog(
            "tiwater.docx.template-migration-choice-catalog/v1",
            true,
            discovery.SourceSha256,
            discovery.BaselineSha256,
            sources,
            targets);
    }

    public static int RunStartDecisions(string[] args)
    {
        if (args.Length < 3) throw new InvalidOperationException("start-template-migration-decisions requires <source.docx> <baseline.docx> <draft.json>");
        Console.WriteLine(JsonSerializer.Serialize(StartDecisionDraft(args[0], args[1], args[2]), Json.CamelCaseOptions));
        return 0;
    }

    public static int RunListDecisionTargets(string[] args)
    {
        if (args.Length < 4) throw new InvalidOperationException("find-template-migration-targets requires <source.docx> <baseline.docx> <draft.json> <branch> [query|-] [offset] [limit]");
        var alignedMapping = string.Equals(args[3], "mapping", StringComparison.Ordinal);
        var branch = alignedMapping
            ? args.Length > 4 && IsTargetedMappingDisposition(args[4])
                ? args[4]
                : throw new InvalidOperationException("template-migration-target-disposition-invalid")
            : args[3];
        var optionalIndex = alignedMapping ? 5 : 4;
        var query = args.Length > optionalIndex && args[optionalIndex] != "-" ? args[optionalIndex] : null;
        var offset = args.Length > optionalIndex + 1 ? int.Parse(args[optionalIndex + 1], System.Globalization.CultureInfo.InvariantCulture) : 0;
        var limit = args.Length > optionalIndex + 2 ? int.Parse(args[optionalIndex + 2], System.Globalization.CultureInfo.InvariantCulture) : 20;
        var result = File.Exists(args[2]) || args[2].EndsWith(".json", StringComparison.OrdinalIgnoreCase)
            ? ListCurrentDecisionTargets(args[0], args[1], args[2], branch, query, offset, limit)
            : ListDecisionTargets(args[0], args[1], args[2] == "-" ? null : args[2], branch, query, offset, limit);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.CamelCaseOptions));
        return 0;
    }

    public static int RunRecordDecision(string[] args)
    {
        if (args.Length < 5) throw new InvalidOperationException("record-template-migration-decision requires <source.docx> <baseline.docx> <draft.json> <branch> <branch arguments>");
        var currentSourceChoiceId = CurrentDecisionSourceId(args[0], args[1], args[2]);
        TemplateMigrationDecisionInput decision = args[3] switch
        {
            "mapping" when args.Length >= 6 && IsMappingDisposition(args[4]) => new TemplateMigrationDecisionInput(
                "mapping", currentSourceChoiceId, args[5] == "-" ? null : args[5], args[4], args.Length > 6 && args[6] != "-" ? args[6] : null),
            "mapping" when args.Length >= 7 => new TemplateMigrationDecisionInput(
                "mapping", args[4], args[6] == "-" ? null : args[6], args[5], args.Length > 7 && args[7] != "-" ? args[7] : null),
            "choice-selection" when args.Length == 5 => new TemplateMigrationDecisionInput(
                "choice-selection", currentSourceChoiceId, args[4]),
            "choice-selection" when args.Length >= 6 => new TemplateMigrationDecisionInput(
                "choice-selection", args[4], args[5]),
            "baseline-clear" when args.Length >= 6 => new TemplateMigrationDecisionInput(
                "baseline-clear", TargetChoiceId: args[4], Mode: args[5]),
            _ => throw new InvalidOperationException("template-migration-decision-arguments-invalid")
        };
        Console.WriteLine(JsonSerializer.Serialize(RecordDecision(args[0], args[1], args[2], decision), Json.CamelCaseOptions));
        return 0;
    }

    public static int RunReviseDecision(string[] args)
    {
        if (args.Length < 6) throw new InvalidOperationException("revise-template-migration-decision requires <source.docx> <baseline.docx> <draft.json> <source-choice-id> <branch> <branch arguments>");
        var decision = args[4] switch
        {
            "mapping" when args.Length >= 7 && IsMappingDisposition(args[5]) => new TemplateMigrationDecisionInput(
                "mapping", args[3], args[6] == "-" ? null : args[6], args[5], args.Length > 7 && args[7] != "-" ? args[7] : null),
            "choice-selection" when args.Length == 6 => new TemplateMigrationDecisionInput(
                "choice-selection", args[3], args[5]),
            _ => throw new InvalidOperationException("template-migration-decision-arguments-invalid")
        };
        Console.WriteLine(JsonSerializer.Serialize(ReviseDecision(args[0], args[1], args[2], decision), Json.CamelCaseOptions));
        return 0;
    }

    public static int RunResolveDecisionDraft(string[] args)
    {
        if (args.Length < 3) throw new InvalidOperationException("resolve-template-migration-decisions requires <source.docx> <baseline.docx> <draft.json>");
        var result = ResolveDecisionDraft(args[0], args[1], args[2]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass || IsClosedReview(result) ? 0 : 1;
    }

    public static TemplateMigrationDecisionProgress StartDecisionDraft(string source, string baseline, string draftFile)
    {
        if (File.Exists(draftFile)) throw new InvalidOperationException("template-migration-decision-draft-already-exists");
        var catalog = ListChoices(source, baseline);
        var draft = new TemplateMigrationDecisionDraft(
            "tiwater.docx.template-migration-decision-draft/v1",
            catalog.SourceSha256,
            catalog.BaselineSha256,
            [],
            [],
            []);
        WriteDecisionDraft(draftFile, draft, overwrite: false);
        return DecisionProgress(catalog, draft);
    }

    public static TemplateMigrationTargetPage ListDecisionTargets(
        string source,
        string baseline,
        string? sourceChoiceId,
        string branch,
        string? query,
        int offset,
        int limit)
        => ListDecisionTargets(source, baseline, sourceChoiceId, branch, query, offset, limit, null);

    private static TemplateMigrationTargetPage ListDecisionTargets(
        string source,
        string baseline,
        string? sourceChoiceId,
        string branch,
        string? query,
        int offset,
        int limit,
        IReadOnlySet<string>? excludedTargetIds)
    {
        if (offset < 0) throw new InvalidOperationException("template-migration-target-offset-invalid");
        if (limit is < 1 or > 100) throw new InvalidOperationException("template-migration-target-limit-invalid");
        var catalog = ListChoices(source, baseline);
        TemplateMigrationChoice? sourceChoice = null;
        if (!string.Equals(branch, "baseline-clear", StringComparison.Ordinal))
        {
            sourceChoice = catalog.Sources.SingleOrDefault(item => string.Equals(item.Id, sourceChoiceId, StringComparison.Ordinal))
                ?? throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale");
        }
        else if (!string.IsNullOrWhiteSpace(sourceChoiceId))
        {
            throw new InvalidOperationException("template-migration-target-source-forbidden");
        }

        IEnumerable<TemplateMigrationChoice> targets = branch switch
        {
            "copy-text" or "retain-target" or "retain-target-label" => catalog.Targets.Where(item =>
                item.AllowedActions?.Contains(PublicActionForDisposition(branch), StringComparer.Ordinal) == true
                && string.Equals(item.Kind, sourceChoice!.Kind, StringComparison.Ordinal)),
            "copy-media" => catalog.Targets.Where(item => item.AllowedActions?.Contains("place-content", StringComparer.Ordinal) == true
                && item.Kind == "media" && sourceChoice!.Kind == "media"),
            "choice-selection" => catalog.Targets.Where(item => item.AllowedActions?.Contains("select-template-option", StringComparer.Ordinal) == true),
            "baseline-clear" => catalog.Targets.Where(item => item.AllowedActions?.Contains("template-cleanup", StringComparer.Ordinal) == true),
            _ => throw new InvalidOperationException("template-migration-target-branch-invalid")
        };
        if (!string.IsNullOrWhiteSpace(query) && query != "-")
        {
            targets = targets.Where(item => ChoiceSearchText(item).Contains(query, StringComparison.OrdinalIgnoreCase));
        }
        if (excludedTargetIds is not null)
        {
            targets = targets.Where(item => !excludedTargetIds.Contains(item.Id));
        }
        var all = targets.OrderBy(item => item.Id, StringComparer.Ordinal).ToList();
        return new TemplateMigrationTargetPage(
            "tiwater.docx.template-migration-target-page/v1",
            true,
            sourceChoice?.Id,
            branch,
            offset,
            limit,
            all.Count,
            all.Skip(offset).Take(limit).ToList());
    }

    public static TemplateMigrationTargetPage ListCurrentDecisionTargets(
        string source,
        string baseline,
        string draftFile,
        string branch,
        string? query,
        int offset,
        int limit)
    {
        var catalog = ListChoices(source, baseline);
        var draft = ReadDecisionDraft(draftFile);
        ValidateDecisionDraftIdentity(catalog, draft);
        ValidateDecisionDraftContent(catalog, draft);
        var sourceChoiceId = string.Equals(branch, "baseline-clear", StringComparison.Ordinal)
            ? null
            : DecisionProgress(catalog, draft).NextSource?.Id
                ?? throw new InvalidOperationException("template-migration-decision-draft-complete");
        var usedTargets = draft.Mappings
            .Where(item => item.TargetChoiceId is not null)
            .Select(item => item.TargetChoiceId!)
            .Concat(draft.ChoiceSelections.Select(item => item.TargetChoiceId))
            .Concat(draft.BaselineClears.Select(item => item.TargetChoiceId))
            .ToHashSet(StringComparer.Ordinal);
        return ListDecisionTargets(source, baseline, sourceChoiceId, branch, query, offset, limit, usedTargets);
    }

    public static TemplateMigrationDecisionProgress RecordDecision(
        string source,
        string baseline,
        string draftFile,
        TemplateMigrationDecisionInput decision)
    {
        var catalog = ListChoices(source, baseline);
        var draft = ReadDecisionDraft(draftFile);
        ValidateDecisionDraftIdentity(catalog, draft);
        ValidateDecisionDraftContent(catalog, draft);
        var mappings = draft.Mappings.ToList();
        var selections = draft.ChoiceSelections.ToList();
        var clears = draft.BaselineClears.ToList();
        if (decision.Branch is "mapping" or "choice-selection")
        {
            mappings.RemoveAll(item => string.Equals(item.SourceChoiceId, decision.SourceChoiceId, StringComparison.Ordinal));
            selections.RemoveAll(item => string.Equals(item.SourceChoiceId, decision.SourceChoiceId, StringComparison.Ordinal));
        }
        var usedSources = mappings.Select(item => item.SourceChoiceId)
            .Concat(selections.Select(item => item.SourceChoiceId))
            .ToHashSet(StringComparer.Ordinal);
        var usedTargets = mappings.Where(item => item.TargetChoiceId is not null).Select(item => item.TargetChoiceId!)
            .Concat(selections.Select(item => item.TargetChoiceId))
            .ToHashSet(StringComparer.Ordinal);

        if (decision.Branch is "mapping" or "choice-selection")
        {
            if (string.IsNullOrWhiteSpace(decision.SourceChoiceId)) throw new InvalidOperationException("template-migration-decision-source-required");
            usedSources.Add(decision.SourceChoiceId);
        }

        if (decision.Branch == "mapping")
        {
            var sourceChoice = catalog.Sources.SingleOrDefault(item => item.Id == decision.SourceChoiceId)
                ?? throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale");
            if (decision.Disposition is not ("copy-text" or "copy-media" or "retain-target" or "retain-target-label" or "out-of-scope" or "review-required"))
            {
                throw new InvalidOperationException("template-migration-decision-disposition-invalid");
            }
            if (decision.Cardinality is not (null or "one" or "all")
                || (decision.Cardinality == "all" && decision.Disposition is not ("out-of-scope" or "review-required")))
            {
                throw new InvalidOperationException("template-migration-decision-cardinality-invalid");
            }
            if (sourceChoice.RequiredCardinality == "all"
                && decision.Disposition == "review-required")
            {
                throw new InvalidOperationException("template-migration-decision-review-group-unsupported");
            }
            if (sourceChoice.RequiredCardinality == "all"
                && (decision.Disposition != "out-of-scope" || decision.Cardinality != "all"))
            {
                throw new InvalidOperationException("template-migration-decision-cardinality-all-required");
            }
            if (decision.Disposition is "out-of-scope" or "review-required")
            {
                if (!string.IsNullOrWhiteSpace(decision.TargetChoiceId)) throw new InvalidOperationException("template-migration-decision-target-forbidden");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(decision.TargetChoiceId)) throw new InvalidOperationException("template-migration-decision-target-required");
                var target = catalog.Targets.SingleOrDefault(item => item.Id == decision.TargetChoiceId)
                    ?? throw new InvalidOperationException("template-migration-decision-target-unknown-or-stale");
                if (target.AllowedActions?.Contains(PublicActionForDisposition(decision.Disposition), StringComparer.Ordinal) != true
                    || !string.Equals(sourceChoice.Kind, target.Kind, StringComparison.Ordinal)
                    || (decision.Disposition == "copy-media" && sourceChoice.Kind != "media")
                    || (decision.Disposition != "copy-media" && sourceChoice.Kind == "media"))
                {
                    throw new InvalidOperationException("template-migration-decision-target-incompatible");
                }
                if (!usedTargets.Add(target.Id)) throw new InvalidOperationException("template-migration-decision-target-duplicate");
            }
            mappings.Add(new TemplateMigrationChoiceMapping(
                sourceChoice.Id,
                decision.Disposition is "out-of-scope" or "review-required" ? null : decision.TargetChoiceId,
                decision.Disposition,
                decision.Cardinality));
        }
        else if (decision.Branch == "choice-selection")
        {
            var sourceChoice = catalog.Sources.SingleOrDefault(item => item.Id == decision.SourceChoiceId)
                ?? throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale");
            var target = catalog.Targets.SingleOrDefault(item => item.Id == decision.TargetChoiceId)
                ?? throw new InvalidOperationException("template-migration-decision-target-unknown-or-stale");
            if (sourceChoice.AllowedActions?.Contains("select-template-option", StringComparer.Ordinal) != true
                || target.AllowedActions?.Contains("select-template-option", StringComparer.Ordinal) != true)
            {
                throw new InvalidOperationException("template-migration-decision-target-incompatible");
            }
            if (!usedTargets.Add(target.Id)) throw new InvalidOperationException("template-migration-decision-target-duplicate");
            selections.Add(new TemplateMigrationChoiceSelectionCandidate(sourceChoice.Id, target.Id));
        }
        else if (decision.Branch == "baseline-clear")
        {
            if (!string.IsNullOrWhiteSpace(decision.SourceChoiceId)) throw new InvalidOperationException("template-migration-decision-source-forbidden");
            var target = catalog.Targets.SingleOrDefault(item => item.Id == decision.TargetChoiceId)
                ?? throw new InvalidOperationException("template-migration-decision-target-unknown-or-stale");
            if (target.AllowedActions?.Contains("template-cleanup", StringComparer.Ordinal) != true
                || decision.Mode is not ("cell" or "row"))
            {
                throw new InvalidOperationException("template-migration-decision-clear-invalid");
            }
            if (usedTargets.Contains(target.Id)) throw new InvalidOperationException("template-migration-decision-clear-target-conflict");
            if (clears.Any(item => item.TargetChoiceId == target.Id)) throw new InvalidOperationException("template-migration-decision-clear-duplicate");
            clears.Add(new TemplateMigrationChoiceClear(target.Id, decision.Mode));
        }
        else
        {
            throw new InvalidOperationException("template-migration-decision-branch-invalid");
        }

        var updated = draft with { Mappings = mappings, ChoiceSelections = selections, BaselineClears = clears };
        ValidateDecisionDraftContent(catalog, updated);
        ValidateDecisionAdmission(source, baseline, catalog, updated);
        WriteDecisionDraft(draftFile, updated, overwrite: true);
        return DecisionProgress(catalog, updated);
    }

    public static TemplateMigrationDecisionProgress ReviseDecision(
        string source,
        string baseline,
        string draftFile,
        TemplateMigrationDecisionInput decision)
    {
        if (decision.Branch is not ("mapping" or "choice-selection"))
        {
            throw new InvalidOperationException("template-migration-decision-revision-branch-invalid");
        }
        if (string.IsNullOrWhiteSpace(decision.SourceChoiceId))
        {
            throw new InvalidOperationException("template-migration-decision-source-required");
        }
        var catalog = ListChoices(source, baseline);
        var draft = ReadDecisionDraft(draftFile);
        ValidateDecisionDraftIdentity(catalog, draft);
        ValidateDecisionDraftContent(catalog, draft);
        var recorded = draft.Mappings.Any(item => string.Equals(item.SourceChoiceId, decision.SourceChoiceId, StringComparison.Ordinal))
            || draft.ChoiceSelections.Any(item => string.Equals(item.SourceChoiceId, decision.SourceChoiceId, StringComparison.Ordinal));
        if (!recorded) throw new InvalidOperationException("template-migration-decision-revision-source-not-recorded");
        return RecordDecision(source, baseline, draftFile, decision);
    }

    private static void ValidateDecisionAdmission(
        string source,
        string baseline,
        TemplateMigrationChoiceCatalog catalog,
        TemplateMigrationDecisionDraft draft)
    {
        var mappings = draft.Mappings
            .Where(item => item.Disposition != "review-required")
            .ToList();
        var usedSources = mappings.Select(item => item.SourceChoiceId)
            .Concat(draft.ChoiceSelections.Select(item => item.SourceChoiceId))
            .ToHashSet(StringComparer.Ordinal);
        mappings.AddRange(catalog.Sources
            .Where(item => !usedSources.Contains(item.Id))
            .Select(item => new TemplateMigrationChoiceMapping(
                item.Id,
                null,
                "out-of-scope",
                item.RequiredCardinality == "all" ? "all" : null)));
        var resolution = ResolveChoices(source, baseline, new TemplateMigrationChoiceCandidate(
            "tiwater.docx.template-migration-choice-candidate/v1",
            mappings,
            draft.ChoiceSelections,
            draft.BaselineClears));
        if (resolution.Pass) return;
        throw new InvalidOperationException(
            resolution.Unresolved.FirstOrDefault()?.Reason
                ?? "template-migration-decision-semantic-admission-failed");
    }

    public static TemplateMigrationMappingDerivation ResolveDecisionDraft(string source, string baseline, string draftFile)
    {
        var catalog = ListChoices(source, baseline);
        var draft = ReadDecisionDraft(draftFile);
        return ResolveDecisionDraft(source, baseline, catalog, draft);
    }

    private static TemplateMigrationMappingDerivation ResolveDecisionDraft(
        string source,
        string baseline,
        TemplateMigrationChoiceCatalog catalog,
        TemplateMigrationDecisionDraft draft)
    {
        ValidateDecisionDraftIdentity(catalog, draft);
        ValidateDecisionDraftContent(catalog, draft);
        if (DecisionProgress(catalog, draft).RemainingSourceCount != 0)
        {
            throw new InvalidOperationException("template-migration-decision-draft-incomplete");
        }
        var reviewMappings = draft.Mappings.Where(item => item.Disposition == "review-required").ToList();
        var determinateMappings = draft.Mappings.Where(item => item.Disposition != "review-required").ToList();
        TemplateMigrationMappingDerivation resolution;
        if (determinateMappings.Count == 0 && draft.ChoiceSelections.Count == 0 && draft.BaselineClears.Count == 0)
        {
            var automatic = DeriveExactTextPlan(source, baseline);
            if (catalog.Sources.Count == 0) return automatic;
            var pendingSourceIds = automatic.Unresolved
                .Where(item => item.SourceObjectId is not null)
                .Select(item => item.SourceObjectId!)
                .ToHashSet(StringComparer.Ordinal);
            resolution = new TemplateMigrationMappingDerivation(
                "tiwater.docx.template-migration-semantic-resolution/v1",
                false,
                automatic.Plan with
                {
                    Mappings = automatic.Plan.Mappings
                        .Where(item => !pendingSourceIds.Contains(item.SourceObjectId))
                        .ToList()
                },
                automatic.Unresolved);
        }
        else
        {
            resolution = ResolveChoices(source, baseline, new TemplateMigrationChoiceCandidate(
                "tiwater.docx.template-migration-choice-candidate/v1",
                determinateMappings,
                draft.ChoiceSelections,
                draft.BaselineClears));
        }
        if (reviewMappings.Count == 0) return resolution;
        var discovery = FindCandidates(source, baseline);
        var sourceSelectors = discovery.RequiredDecisions.ToDictionary(
            item => ChoiceId("source", discovery.SourceSha256, item.Source.Selector
                ?? throw new InvalidOperationException("template-migration-choice-source-selector-missing")),
            item => item.Source.Selector!,
            StringComparer.Ordinal);
        var reviews = reviewMappings.Select(item => new TemplateMigrationSemanticCandidateMapping(
            sourceSelectors.TryGetValue(item.SourceChoiceId, out var selector)
                ? selector
                : throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale"),
            null,
            "review-required",
            item.Cardinality)).ToList();
        return CloseReviews(source, baseline, resolution, new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v5",
            reviews));
    }

    private static TemplateMigrationDecisionProgress DecisionProgress(
        TemplateMigrationChoiceCatalog catalog,
        TemplateMigrationDecisionDraft draft)
    {
        var usedSources = draft.Mappings.Select(item => item.SourceChoiceId)
            .Concat(draft.ChoiceSelections.Select(item => item.SourceChoiceId))
            .ToHashSet(StringComparer.Ordinal);
        var remaining = catalog.Sources.Where(item => !usedSources.Contains(item.Id)).ToList();
        return new TemplateMigrationDecisionProgress(
            "tiwater.docx.template-migration-decision-progress/v1",
            true,
            usedSources.Count,
            remaining.Count,
            remaining.FirstOrDefault());
    }

    private static string? CurrentDecisionSourceId(string source, string baseline, string draftFile)
    {
        var catalog = ListChoices(source, baseline);
        var draft = ReadDecisionDraft(draftFile);
        ValidateDecisionDraftIdentity(catalog, draft);
        ValidateDecisionDraftContent(catalog, draft);
        return DecisionProgress(catalog, draft).NextSource?.Id;
    }

    private static bool IsMappingDisposition(string value)
        => value is "copy-text" or "copy-media" or "retain-target" or "retain-target-label" or "out-of-scope" or "review-required";

    private static bool IsTargetedMappingDisposition(string value)
        => value is "copy-text" or "copy-media" or "retain-target" or "retain-target-label";

    private static string PublicActionForDisposition(string value)
        => value switch
        {
            "copy-text" or "copy-media" => "place-content",
            "retain-target" => "keep-template-content",
            "retain-target-label" => "keep-template-label",
            _ => throw new InvalidOperationException("template-migration-decision-disposition-invalid")
        };

    private static string ChoiceSearchText(TemplateMigrationChoice choice)
        => string.Join("\n", new[]
        {
            choice.Text,
            choice.Context?.PreviousText,
            choice.Context?.NextText,
            choice.Context?.SameRowTexts is null ? null : string.Join("\n", choice.Context.SameRowTexts)
        }.Where(item => !string.IsNullOrWhiteSpace(item))!);

    private static void ValidateDecisionDraftIdentity(TemplateMigrationChoiceCatalog catalog, TemplateMigrationDecisionDraft draft)
    {
        if (draft.Schema != "tiwater.docx.template-migration-decision-draft/v1") throw new InvalidOperationException("template-migration-decision-draft-schema-invalid");
        if (!string.Equals(draft.SourceSha256, catalog.SourceSha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("template-migration-decision-draft-source-stale");
        if (!string.Equals(draft.BaselineSha256, catalog.BaselineSha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("template-migration-decision-draft-baseline-stale");
    }

    private static void ValidateDecisionDraftContent(TemplateMigrationChoiceCatalog catalog, TemplateMigrationDecisionDraft draft)
    {
        var determinateMappings = draft.Mappings.Where(item => item.Disposition != "review-required").ToList();
        if (determinateMappings.Count != 0 || draft.ChoiceSelections.Count != 0 || draft.BaselineClears.Count != 0)
        {
            ValidateChoiceCandidate(new TemplateMigrationChoiceCandidate(
                "tiwater.docx.template-migration-choice-candidate/v1",
                determinateMappings,
                draft.ChoiceSelections,
                draft.BaselineClears));
        }
        var sources = catalog.Sources.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var targets = catalog.Targets.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var usedSources = new HashSet<string>(StringComparer.Ordinal);
        var usedTargets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var mapping in draft.Mappings)
        {
            if (!usedSources.Add(mapping.SourceChoiceId)) throw new InvalidOperationException("template-migration-decision-source-duplicate");
            if (!sources.TryGetValue(mapping.SourceChoiceId, out var source)) throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale");
            if (mapping.Disposition is not ("copy-text" or "copy-media" or "retain-target" or "retain-target-label" or "out-of-scope" or "review-required"))
            {
                throw new InvalidOperationException("template-migration-decision-disposition-invalid");
            }
            if (mapping.Disposition is "out-of-scope" or "review-required")
            {
                if (mapping.TargetChoiceId is not null) throw new InvalidOperationException("template-migration-decision-target-forbidden");
                if (mapping.Disposition == "review-required" && source.RequiredCardinality == "all") throw new InvalidOperationException("template-migration-decision-review-group-unsupported");
                if (source.RequiredCardinality == "all" && mapping.Cardinality != "all") throw new InvalidOperationException("template-migration-decision-cardinality-all-required");
                continue;
            }
            if (mapping.TargetChoiceId is null || !targets.TryGetValue(mapping.TargetChoiceId, out var target)) throw new InvalidOperationException("template-migration-decision-target-unknown-or-stale");
            if (target.AllowedActions?.Contains(PublicActionForDisposition(mapping.Disposition), StringComparer.Ordinal) != true
                || !string.Equals(source.Kind, target.Kind, StringComparison.Ordinal)
                || (mapping.Disposition == "copy-media" && source.Kind != "media")
                || (mapping.Disposition != "copy-media" && source.Kind == "media"))
            {
                throw new InvalidOperationException("template-migration-decision-target-incompatible");
            }
            if (!usedTargets.Add(target.Id)) throw new InvalidOperationException("template-migration-decision-target-duplicate");
        }
        foreach (var selection in draft.ChoiceSelections)
        {
            if (!usedSources.Add(selection.SourceChoiceId)) throw new InvalidOperationException("template-migration-decision-source-duplicate");
            if (!sources.TryGetValue(selection.SourceChoiceId, out var source)) throw new InvalidOperationException("template-migration-decision-source-unknown-or-stale");
            if (!targets.TryGetValue(selection.TargetChoiceId, out var target)) throw new InvalidOperationException("template-migration-decision-target-unknown-or-stale");
            if (source.AllowedActions?.Contains("select-template-option", StringComparer.Ordinal) != true
                || target.AllowedActions?.Contains("select-template-option", StringComparer.Ordinal) != true)
            {
                throw new InvalidOperationException("template-migration-decision-target-incompatible");
            }
            if (!usedTargets.Add(target.Id)) throw new InvalidOperationException("template-migration-decision-target-duplicate");
        }
        foreach (var clear in draft.BaselineClears)
        {
            if (!targets.TryGetValue(clear.TargetChoiceId, out var target)
                || target.AllowedActions?.Contains("template-cleanup", StringComparer.Ordinal) != true)
            {
                throw new InvalidOperationException("template-migration-decision-clear-invalid");
            }
            if (!usedTargets.Add(target.Id)) throw new InvalidOperationException("template-migration-decision-clear-target-conflict");
        }
    }

    private static TemplateMigrationDecisionDraft ReadDecisionDraft(string file)
    {
        if (!File.Exists(file)) throw new InvalidOperationException("template-migration-decision-draft-missing");
        using var document = JsonDocument.Parse(File.ReadAllText(file));
        var root = document.RootElement;
        RequireOnlyFields(root, new HashSet<string>(["schema", "sourceSha256", "baselineSha256", "mappings", "choiceSelections", "baselineClears"], StringComparer.Ordinal), "template-migration-decision-draft");
        foreach (var field in new[] { "schema", "sourceSha256", "baselineSha256" })
        {
            if (!root.TryGetProperty(field, out var value) || value.ValueKind != JsonValueKind.String) throw new InvalidOperationException($"template-migration-decision-draft-{field}-invalid");
        }
        foreach (var field in new[] { "mappings", "choiceSelections", "baselineClears" })
        {
            if (!root.TryGetProperty(field, out var value) || value.ValueKind != JsonValueKind.Array) throw new InvalidOperationException($"template-migration-decision-draft-{field}-invalid");
        }
        foreach (var mapping in root.GetProperty("mappings").EnumerateArray())
        {
            RequireOnlyFields(mapping, new HashSet<string>(["sourceChoiceId", "targetChoiceId", "disposition", "cardinality"], StringComparer.Ordinal), "template-migration-decision-draft-mapping");
        }
        foreach (var selection in root.GetProperty("choiceSelections").EnumerateArray())
        {
            RequireOnlyFields(selection, new HashSet<string>(["sourceChoiceId", "targetChoiceId"], StringComparer.Ordinal), "template-migration-decision-draft-selection");
        }
        foreach (var clear in root.GetProperty("baselineClears").EnumerateArray())
        {
            RequireOnlyFields(clear, new HashSet<string>(["targetChoiceId", "mode"], StringComparer.Ordinal), "template-migration-decision-draft-clear");
        }
        var draft = JsonSerializer.Deserialize<TemplateMigrationDecisionDraft>(root.GetRawText(), Json.CamelCaseOptions)
            ?? throw new InvalidOperationException("template-migration-decision-draft-invalid");
        if (draft.Mappings is null || draft.ChoiceSelections is null || draft.BaselineClears is null) throw new InvalidOperationException("template-migration-decision-draft-invalid");
        return draft;
    }

    private static void WriteDecisionDraft(string file, TemplateMigrationDecisionDraft draft, bool overwrite)
    {
        var fullPath = Path.GetFullPath(file);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var temporary = fullPath + ".tmp-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(temporary, JsonSerializer.Serialize(draft, Json.CamelCaseOptions));
        try
        {
            if (!overwrite && File.Exists(fullPath)) throw new InvalidOperationException("template-migration-decision-draft-already-exists");
            File.Move(temporary, fullPath, overwrite);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    public static int RunResolveChoices(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("resolve-template-migration-choices requires <source.docx> <baseline.docx> <choices.json>");
        }
        var candidate = ReadChoiceCandidate(args[2]);
        var result = ResolveChoices(args[0], args[1], candidate);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    public static TemplateMigrationMappingDerivation ResolveChoices(
        string source,
        string baseline,
        TemplateMigrationChoiceCandidate candidate)
    {
        ValidateChoiceCandidate(candidate);
        var discovery = FindCandidates(source, baseline);
        var sourceSelectors = discovery.RequiredDecisions.ToDictionary(
            item => ChoiceId("source", discovery.SourceSha256, item.Source.Selector
                ?? throw new InvalidOperationException("template-migration-choice-source-selector-missing")),
            item => item.Source.Selector!,
            StringComparer.Ordinal);
        var targetSelectors = discovery.AvailableTargets
            .SelectMany(item => new[] { item }.Concat(item.Context?.SelectableChildren ?? []))
            .GroupBy(item => ChoiceId("target", discovery.BaselineSha256, item.Selector
                ?? throw new InvalidOperationException("template-migration-choice-target-selector-missing")), StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First().Selector!, StringComparer.Ordinal);

        var usedSources = new HashSet<string>(StringComparer.Ordinal);
        var mappings = new List<TemplateMigrationSemanticCandidateMapping>();
        foreach (var mapping in candidate.Mappings)
        {
            if (!usedSources.Add(mapping.SourceChoiceId)) throw new InvalidOperationException("template-migration-choice-source-duplicate");
            if (!sourceSelectors.TryGetValue(mapping.SourceChoiceId, out var sourceSelector)) throw new InvalidOperationException("template-migration-choice-source-unknown-or-stale");
            TemplateMigrationSemanticSelector? targetSelector = null;
            if (mapping.TargetChoiceId is not null
                && !targetSelectors.TryGetValue(mapping.TargetChoiceId, out targetSelector))
            {
                throw new InvalidOperationException("template-migration-choice-target-unknown-or-stale");
            }
            mappings.Add(new TemplateMigrationSemanticCandidateMapping(
                sourceSelector,
                targetSelector,
                mapping.Disposition,
                mapping.Cardinality));
        }

        var selections = new List<TemplateMigrationSemanticCandidateChoiceSelection>();
        foreach (var selection in candidate.ChoiceSelections ?? [])
        {
            if (!usedSources.Add(selection.SourceChoiceId)) throw new InvalidOperationException("template-migration-choice-source-duplicate");
            if (!sourceSelectors.TryGetValue(selection.SourceChoiceId, out var sourceSelector)) throw new InvalidOperationException("template-migration-choice-source-unknown-or-stale");
            if (!targetSelectors.TryGetValue(selection.TargetChoiceId, out var targetSelector)) throw new InvalidOperationException("template-migration-choice-target-unknown-or-stale");
            selections.Add(new TemplateMigrationSemanticCandidateChoiceSelection(sourceSelector, targetSelector));
        }

        var clearedTargets = new HashSet<string>(StringComparer.Ordinal);
        var clears = new List<TemplateMigrationSemanticCandidateBaselineClear>();
        foreach (var clear in candidate.BaselineClears ?? [])
        {
            if (!clearedTargets.Add(clear.TargetChoiceId)) throw new InvalidOperationException("template-migration-choice-clear-target-duplicate");
            if (!targetSelectors.TryGetValue(clear.TargetChoiceId, out var targetSelector)) throw new InvalidOperationException("template-migration-choice-target-unknown-or-stale");
            clears.Add(new TemplateMigrationSemanticCandidateBaselineClear(targetSelector, clear.Mode));
        }

        var semanticSchema = mappings.Any(item => UsesEmptyTextState(item.Source) || UsesEmptyTextState(item.Baseline))
            || selections.Any(item => UsesEmptyTextState(item.SourceMember) || UsesEmptyTextState(item.BaselineLabel))
            || clears.Any(item => UsesEmptyTextState(item.Baseline))
            ? "tiwater.docx.template-migration-semantic-candidate/v6"
            : "tiwater.docx.template-migration-semantic-candidate/v5";
        var semanticCandidate = new TemplateMigrationSemanticCandidate(
            semanticSchema,
            mappings,
            ChoiceSelections: selections,
            BaselineClears: clears);
        ValidateSemanticCandidate(semanticCandidate);
        var analysis = Analyze(source, baseline);
        return ResolveSemanticCandidate(
            source,
            baseline,
            semanticCandidate,
            analysis,
            DeriveExactTextPlan(analysis));
    }

    public static int RunMigrateTemplate(string[] args)
    {
        if (args.Length < 4)
        {
            throw new InvalidOperationException("migrate-template requires <source.docx> <baseline.docx> <choices.json> <output.docx>");
        }
        var receipt = MigrateTemplate(args[0], args[1], ReadBusinessChoiceBatch(args[2]), args[3]);
        Console.WriteLine(JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
        return receipt.Pass ? 0 : 1;
    }

    public static int RunVerifyTemplateMigration(string[] args)
    {
        if (args.Length < 4)
        {
            throw new InvalidOperationException("verify-template-migration requires <source.docx> <baseline.docx> <choices.json> <output.docx>");
        }
        var receipt = VerifyTemplateMigration(args[0], args[1], ReadBusinessChoiceBatch(args[2]), args[3]);
        Console.WriteLine(JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
        return receipt.Pass ? 0 : 1;
    }

    public static TemplateMigrationExecutionReceipt MigrateTemplate(
        string source,
        string baseline,
        TemplateMigrationBusinessChoiceBatch choices,
        string output)
    {
        var version = typeof(TemplateMigration).Assembly.GetName().Version?.ToString() ?? "unknown";
        var resolution = ResolveBusinessChoices(source, baseline, choices);
        if (!resolution.Pass)
        {
            var review = IsClosedReview(resolution);
            if (review)
            {
                return MigrateReviewTemplate(source, baseline, output, version, resolution);
            }
            return new TemplateMigrationExecutionReceipt(
                "tiwater.docx.template-migration-execution/v1", version,
                "failed", false, false, false,
                null, null, resolution, null, null, null, null, resolution.Unresolved);
        }

        var build = BuildOperations(source, baseline, resolution.Plan);
        if (!build.Pass)
        {
            return new TemplateMigrationExecutionReceipt(
                "tiwater.docx.template-migration-execution/v1", version,
                "failed", false, false, false, null, null, resolution, build, null, null, null, build.Failures);
        }

        var outputPath = Path.GetFullPath(output);
        var planPath = MigrationPlanPath(outputPath);
        if (File.Exists(outputPath)) throw new InvalidOperationException("template-migration-output-already-exists");
        if (File.Exists(planPath)) throw new InvalidOperationException("template-migration-plan-already-exists");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory());
        var temporaryOutput = Path.Combine(Path.GetDirectoryName(outputPath)!, $".{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.verified");
        var temporaryPlan = Path.Combine(Path.GetDirectoryName(outputPath)!, $".{Path.GetFileName(planPath)}.{Guid.NewGuid():N}.pending");
        TemplateMigrationApplyResult? apply = null;
        TemplateMigrationOutputValidation? validation = null;
        try
        {
            File.WriteAllText(temporaryPlan, JsonSerializer.Serialize(resolution.Plan, Json.Options));
            apply = Apply(source, baseline, resolution.Plan, temporaryOutput);
            if (!apply.Pass)
            {
                return new TemplateMigrationExecutionReceipt(
                    "tiwater.docx.template-migration-execution/v1", version,
                    "failed", false, false, false, null, null, resolution, build, apply, null, null,
                    [.. apply.Build.Failures, .. apply.MediaFailures, .. (apply.Readback?.Failures ?? [])]);
            }
            validation = ValidateOutput(source, baseline, temporaryPlan, temporaryOutput, resolution.Plan);
            if (!validation.Pass)
            {
                return new TemplateMigrationExecutionReceipt(
                    "tiwater.docx.template-migration-execution/v1", version,
                    "failed", false, false, false, null, null, resolution, build, apply, null, validation, validation.Failures);
            }
            File.Move(temporaryOutput, outputPath);
            File.Move(temporaryPlan, planPath);
            validation = ValidateOutput(source, baseline, planPath, outputPath, resolution.Plan);
            return new TemplateMigrationExecutionReceipt(
                "tiwater.docx.template-migration-execution/v1", version,
                validation.Pass ? "pass" : "failed", validation.Pass, false, validation.Pass,
                validation.Pass ? outputPath : null, planPath, resolution, build, apply, null, validation, validation.Failures);
        }
        finally
        {
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
            if (File.Exists(temporaryPlan)) File.Delete(temporaryPlan);
        }
    }

    private static TemplateMigrationExecutionReceipt MigrateReviewTemplate(
        string source,
        string baseline,
        string output,
        string version,
        TemplateMigrationMappingDerivation resolution)
    {
        var build = BuildOperations(source, baseline, resolution.Plan);
        if (build.Failures.Count != 0)
        {
            return new TemplateMigrationExecutionReceipt(
                "tiwater.docx.template-migration-execution/v1", version,
                "failed", false, true, false, null, null, resolution, build, null, null, null,
                [.. resolution.Unresolved, .. build.Failures]);
        }

        var outputPath = Path.GetFullPath(output);
        var planPath = MigrationPlanPath(outputPath);
        if (File.Exists(outputPath)) throw new InvalidOperationException("template-migration-output-already-exists");
        if (File.Exists(planPath)) throw new InvalidOperationException("template-migration-plan-already-exists");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory());
        var temporaryOutput = Path.Combine(Path.GetDirectoryName(outputPath)!, $".{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.review");
        var temporaryPlan = Path.Combine(Path.GetDirectoryName(outputPath)!, $".{Path.GetFileName(planPath)}.{Guid.NewGuid():N}.pending");
        try
        {
            File.WriteAllText(temporaryPlan, JsonSerializer.Serialize(resolution.Plan, Json.Options));
            var preview = Preview(source, baseline, resolution.Plan, temporaryOutput);
            if (!preview.OutputVerified)
            {
                return new TemplateMigrationExecutionReceipt(
                    "tiwater.docx.template-migration-execution/v1", version,
                    "failed", false, true, false, null, null, resolution, build, null, preview, null,
                    [.. resolution.Unresolved, .. preview.Build.Failures, .. preview.MediaFailures, .. (preview.Readback?.Failures ?? [])]);
            }
            File.Move(temporaryOutput, outputPath);
            File.Move(temporaryPlan, planPath);
            return new TemplateMigrationExecutionReceipt(
                "tiwater.docx.template-migration-execution/v1", version,
                "review-required", false, true, true, outputPath, planPath,
                resolution, build, null, preview, null, resolution.Unresolved);
        }
        finally
        {
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
            if (File.Exists(temporaryPlan)) File.Delete(temporaryPlan);
        }
    }

    public static TemplateMigrationVerificationReceipt VerifyTemplateMigration(
        string source,
        string baseline,
        TemplateMigrationBusinessChoiceBatch choices,
        string output)
    {
        var version = typeof(TemplateMigration).Assembly.GetName().Version?.ToString() ?? "unknown";
        var outputPath = Path.GetFullPath(output);
        var planPath = MigrationPlanPath(outputPath);
        var resolution = ResolveBusinessChoices(source, baseline, choices);
        var review = IsClosedReview(resolution);
        var failures = new List<TemplateMigrationPlanFailure>();
        failures.AddRange(resolution.Unresolved);
        if ((!resolution.Pass && !review) || !File.Exists(outputPath) || !File.Exists(planPath))
        {
            if (!File.Exists(outputPath)) failures.Add(new TemplateMigrationPlanFailure("template-migration-output-missing"));
            if (!File.Exists(planPath)) failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-missing"));
            return new TemplateMigrationVerificationReceipt(
                "tiwater.docx.template-migration-verification/v1", version, "failed", false, review, false,
                outputPath, File.Exists(planPath) ? planPath : null, resolution, null, failures);
        }

        var storedPlan = JsonSerializer.Deserialize<TemplateMigrationPlan>(File.ReadAllText(planPath), Json.Options)
            ?? throw new InvalidOperationException("template-migration-plan-invalid");
        if (!string.Equals(HashCanonical(storedPlan), HashCanonical(resolution.Plan), StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-choice-mismatch"));
            return new TemplateMigrationVerificationReceipt(
                "tiwater.docx.template-migration-verification/v1", version, "failed", false, review, false,
                outputPath, planPath, resolution, null, failures);
        }
        var validation = ValidateOutput(source, baseline, planPath, outputPath, resolution.Plan);
        failures.AddRange(validation.Failures);
        var outputVerified = review
            ? validation.Build.Failures.Count == 0 && validation.Readback.Pass
            : validation.Pass;
        return new TemplateMigrationVerificationReceipt(
            "tiwater.docx.template-migration-verification/v1", version,
            outputVerified ? (review ? "review-required" : "pass") : "failed",
            !review && outputVerified, review, outputVerified,
            outputPath, planPath, resolution, validation, failures);
    }

    public static TemplateMigrationMappingDerivation ResolveBusinessChoices(
        string source,
        string baseline,
        TemplateMigrationBusinessChoiceBatch batch)
    {
        ValidateBusinessChoiceBatch(batch);
        var catalog = ListChoices(source, baseline);
        var sources = catalog.Sources.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var targets = catalog.Targets.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var usedSources = new HashSet<string>(StringComparer.Ordinal);
        var mappings = new List<TemplateMigrationChoiceMapping>();
        var selections = new List<TemplateMigrationChoiceSelectionCandidate>();
        foreach (var choice in batch.Choices)
        {
            if (!usedSources.Add(choice.SourceChoiceId)) throw new InvalidOperationException("template-migration-business-source-duplicate");
            if (!sources.TryGetValue(choice.SourceChoiceId, out var sourceChoice)) throw new InvalidOperationException("template-migration-business-source-unknown-or-stale");
            TemplateMigrationChoice? targetChoice = null;
            if (choice.TargetChoiceId is not null && !targets.TryGetValue(choice.TargetChoiceId, out targetChoice))
                throw new InvalidOperationException("template-migration-business-target-unknown-or-stale");

            switch (choice.Action)
            {
                case "place-content":
                    RequireBusinessTarget(targetChoice, choice.Action);
                    mappings.Add(new TemplateMigrationChoiceMapping(
                        sourceChoice.Id, targetChoice!.Id,
                        sourceChoice.Kind == "media" ? "copy-media" : "copy-text", choice.Cardinality));
                    break;
                case "keep-template-content":
                    RequireBusinessTarget(targetChoice, choice.Action);
                    mappings.Add(new TemplateMigrationChoiceMapping(sourceChoice.Id, targetChoice!.Id, "retain-target", choice.Cardinality));
                    break;
                case "keep-template-label":
                    RequireBusinessTarget(targetChoice, choice.Action);
                    mappings.Add(new TemplateMigrationChoiceMapping(sourceChoice.Id, targetChoice!.Id, "retain-target-label", choice.Cardinality));
                    break;
                case "select-template-option":
                    RequireBusinessTarget(targetChoice, choice.Action);
                    selections.Add(new TemplateMigrationChoiceSelectionCandidate(sourceChoice.Id, targetChoice!.Id));
                    break;
                case "exclude-source":
                    if (targetChoice is not null) throw new InvalidOperationException("template-migration-business-target-forbidden");
                    mappings.Add(new TemplateMigrationChoiceMapping(sourceChoice.Id, null, "out-of-scope", choice.Cardinality));
                    break;
                case "review-source":
                    if (targetChoice is not null) throw new InvalidOperationException("template-migration-business-target-forbidden");
                    mappings.Add(new TemplateMigrationChoiceMapping(sourceChoice.Id, null, "review-required", choice.Cardinality));
                    break;
                default:
                    throw new InvalidOperationException("template-migration-business-action-invalid");
            }
        }
        if (!usedSources.SetEquals(sources.Keys)) throw new InvalidOperationException("template-migration-business-choice-set-incomplete");

        var clears = new List<TemplateMigrationChoiceClear>();
        var clearedTargets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var cleanup in batch.TemplateCleanup ?? [])
        {
            if (!clearedTargets.Add(cleanup.TargetChoiceId)) throw new InvalidOperationException("template-migration-business-cleanup-duplicate");
            if (!targets.TryGetValue(cleanup.TargetChoiceId, out var target)) throw new InvalidOperationException("template-migration-business-target-unknown-or-stale");
            RequireBusinessTarget(target, "template-cleanup");
            clears.Add(new TemplateMigrationChoiceClear(target.Id, cleanup.Scope));
        }

        var draft = new TemplateMigrationDecisionDraft(
            "tiwater.docx.template-migration-decision-draft/v1",
            catalog.SourceSha256, catalog.BaselineSha256, mappings, selections, clears);
        return ResolveDecisionDraft(source, baseline, catalog, draft);
    }

    private static void RequireBusinessTarget(TemplateMigrationChoice? target, string use)
    {
        if (target is null) throw new InvalidOperationException("template-migration-business-target-required");
        if (!(target.AllowedActions ?? []).Contains(use, StringComparer.Ordinal))
            throw new InvalidOperationException("template-migration-business-target-incompatible");
    }

    private static string MigrationPlanPath(string outputPath) => outputPath + ".migration-plan.json";

    private static bool UsesEmptyTextState(TemplateMigrationSemanticSelector? selector)
        => !string.IsNullOrWhiteSpace(selector?.TextState);

    private static TemplateMigrationChoice ToChoice(
        string id,
        TemplateMigrationSemanticObservation observation,
        int count = 1,
        string? requiredCardinality = null,
        IReadOnlyList<string>? allowedActions = null)
        => new(
            id,
            observation.Kind,
            observation.Scope,
            observation.Text,
            count,
            requiredCardinality,
            observation.Context is null
                ? null
                : new TemplateMigrationSemanticContext(
                    observation.Context.PreviousText,
                    observation.Context.NextText,
                    observation.Context.SameRowTexts),
            allowedActions);

    private static string ChoiceId(string role, string documentSha256, TemplateMigrationSemanticSelector selector)
        => $"{role}-{HashCanonical(new { role, documentSha256, selector }).ToLowerInvariant()[..20]}";

    private static void EnsureUniqueChoiceIds(IReadOnlyList<TemplateMigrationChoice> choices, string role)
    {
        if (choices.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != choices.Count)
        {
            throw new InvalidOperationException($"template-migration-choice-{role}-identity-collision");
        }
    }

    /// <summary>
    /// Finds mechanical source/baseline observation pairs without producing a
    /// migration plan or deciding any semantic disposition.
    /// </summary>
    public static TemplateMigrationCandidateDiscovery FindCandidates(string source, string baseline)
    {
        var analysis = Analyze(source, baseline);
        var legacy = DeriveAnchorGapPlan(analysis);
        var automaticPending = legacy.Unresolved
            .Where(item => !string.IsNullOrWhiteSpace(item.SourceObjectId))
            .GroupBy(item => item.SourceObjectId!, StringComparer.Ordinal)
            .Select(group => group.FirstOrDefault(item => item.Reason == "template-migration-anchor-gap-candidate-review-required")
                ?? group.First())
            .ToList();
        if (automaticPending.Any(item => item.Source is null || string.IsNullOrWhiteSpace(item.SourceObjectId)))
        {
            throw new InvalidOperationException("template-migration-candidate-source-observation-missing");
        }
        var pendingWithoutSelectorIds = automaticPending
            .Where(item => item.Source!.Selector is null)
            .Select(item => item.SourceObjectId!)
            .ToHashSet(StringComparer.Ordinal);
        var sourceById = analysis.Source.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var claimedSourceIds = legacy.Plan.Mappings
            .Select(mapping => mapping.SourceObjectId)
            .ToHashSet(StringComparer.Ordinal);
        var coveredSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var decisions = new List<TemplateMigrationRequiredDecision>();
        foreach (var item in automaticPending)
        {
            if (coveredSourceIds.Contains(item.SourceObjectId!)) continue;
            if (item.Source!.Selector is null)
            {
                if (!sourceById.TryGetValue(item.SourceObjectId!, out var sourceObject))
                {
                    throw new InvalidOperationException("template-migration-candidate-source-selector-missing");
                }
                var group = SemanticSelectorCandidates(analysis.Source.Objects, sourceObject)
                    .Select(selector => (Selector: selector, Selected: ResolveSelector(analysis.Source.Objects, selector)))
                    .Where(candidate => candidate.Selected.Count > 1
                        && candidate.Selected.All(selected => pendingWithoutSelectorIds.Contains(selected.Id)
                            && !coveredSourceIds.Contains(selected.Id)))
                    .OrderBy(candidate => candidate.Selected.Count)
                    .FirstOrDefault();
                if (group.Selector is null)
                {
                    throw new InvalidOperationException("template-migration-candidate-source-selector-missing");
                }
                foreach (var selectedItem in group.Selected) coveredSourceIds.Add(selectedItem.Id);
                decisions.Add(new TemplateMigrationRequiredDecision(
                    new TemplateMigrationSemanticObservation(
                        sourceObject.Kind,
                        sourceObject.Scope,
                        sourceObject.Text,
                        group.Selector),
                    group.Selected.Count,
                    "all"));
                continue;
            }
            coveredSourceIds.Add(item.SourceObjectId!);
            decisions.Add(new TemplateMigrationRequiredDecision(
                ObserveContextualRegion(
                    analysis.Source.Objects,
                    sourceById[item.SourceObjectId!],
                    claimedSourceIds)));
        }
        return new TemplateMigrationCandidateDiscovery(
            "tiwater.docx.template-migration-candidate-discovery/v5",
            true,
            legacy.Plan.SourceSha256,
            legacy.Plan.BaselineSha256,
            decisions,
            ObserveAvailableTargets(analysis, legacy.Plan));
    }

    /// <summary>
    /// Derives semantic candidates only when two consecutive
    /// exact-text anchors enclose equally sized paragraph gaps in the same
    /// document scope. It never turns structural adjacency into an operation.
    /// </summary>
    public static TemplateMigrationMappingDerivation DeriveAnchorGapPlan(string source, string baseline)
    {
        var analysis = Analyze(source, baseline);
        return DeriveAnchorGapPlan(analysis);
    }

    private static TemplateMigrationMappingDerivation DeriveAnchorGapPlan(TemplateMigrationAnalysis analysis)
    {
        var exact = DeriveExactTextPlan(analysis);
        var mappings = exact.Plan.Mappings.ToDictionary(mapping => mapping.SourceObjectId, StringComparer.Ordinal);
        var pendingSourceIds = exact.Unresolved
            .Where(item => !string.IsNullOrWhiteSpace(item.SourceObjectId))
            .Select(item => item.SourceObjectId!)
            .ToHashSet(StringComparer.Ordinal);
        var candidates = new List<(TemplateMigrationObject Source, TemplateMigrationObject Baseline)>();
        foreach (var scope in analysis.Source.Objects.Where(item => item.Kind == "paragraph" && IsContentBearing(item)).Select(item => item.Scope).Distinct(StringComparer.Ordinal))
        {
            var sourceParagraphs = analysis.Source.Objects.Where(item => item.Kind == "paragraph" && item.Scope == scope && IsContentBearing(item)).ToList();
            var baselineParagraphs = analysis.Baseline.Objects.Where(item => item.Kind == "paragraph" && item.Scope == scope && IsContentBearing(item)).ToList();
            candidates.AddRange(FindEqualAnchorGapCandidates(sourceParagraphs, baselineParagraphs, mappings, pendingSourceIds));
        }

        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v1",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            mappings.Values.OrderBy(mapping => mapping.SourceObjectId, StringComparer.Ordinal).ToList());
        var unresolved = new List<TemplateMigrationPlanFailure>(exact.Unresolved);
        foreach (var candidate in candidates)
        {
            unresolved.Add(new TemplateMigrationPlanFailure(
                "template-migration-anchor-gap-candidate-review-required",
                candidate.Source.Id,
                candidate.Baseline.Id,
                Source: ObserveForSemanticDecision(candidate.Source),
                Baseline: ObserveForSemanticDecision(candidate.Baseline)));
        }
        return new TemplateMigrationMappingDerivation(
            "tiwater.docx.template-migration-anchor-gap-plan/v1",
            true,
            plan,
            unresolved,
            ObserveUnclaimedBaseline(analysis, plan));
    }

    private static TemplateMigrationSemanticObservation ObserveForSemanticDecision(TemplateMigrationObject item)
        => new(item.Kind, item.Scope, item.Text, item.Selector);

    private static IReadOnlyList<TemplateMigrationSemanticObservation> ObserveUnclaimedBaseline(
        TemplateMigrationAnalysis analysis,
        TemplateMigrationPlan plan)
    {
        var claimed = plan.Mappings
            .Where(mapping => !string.IsNullOrWhiteSpace(mapping.BaselineObjectId))
            .Select(mapping => mapping.BaselineObjectId!)
            .ToHashSet(StringComparer.Ordinal);
        var unclaimedContent = analysis.Baseline.Objects
            .Where(IsContentBearing)
            .Where(item => !claimed.Contains(item.Id))
            .ToList();
        var unclaimedCellIds = unclaimedContent
            .Where(item => item.Kind == "table-cell")
            .Select(item => item.Id)
            .ToHashSet(StringComparer.Ordinal);
        return unclaimedContent
            .Concat(analysis.Baseline.Objects.Where(item =>
                item.Kind == "run"
                && item.Selector is not null
                && !string.IsNullOrWhiteSpace(item.Text)
                && item.ParentId is not null
                && unclaimedCellIds.Contains(item.ParentId)))
            .OrderBy(item => item.Id, StringComparer.Ordinal)
            .Select(ObserveForSemanticDecision)
            .ToList();
    }

    private static IReadOnlyList<TemplateMigrationSemanticObservation> ObserveAvailableTargets(
        TemplateMigrationAnalysis analysis,
        TemplateMigrationPlan plan)
    {
        var claimed = plan.Mappings
            .Where(mapping => !string.IsNullOrWhiteSpace(mapping.BaselineObjectId))
            .Select(mapping => mapping.BaselineObjectId!)
            .ToHashSet(StringComparer.Ordinal);
        return analysis.Baseline.Objects
            .Where(item => item.Kind is "paragraph" or "table-cell" or "media")
            .Where(item => item.Selector is not null && !claimed.Contains(item.Id))
            .OrderBy(item => item.Id, StringComparer.Ordinal)
            .Select(item => ObserveContextualRegion(analysis.Baseline.Objects, item, claimed))
            .ToList();
    }

    private static TemplateMigrationSemanticObservation ObserveContextualRegion(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject item,
        IReadOnlySet<string> claimed)
    {
        var siblings = objects
            .Where(candidate => candidate.Kind == item.Kind
                && candidate.Scope == item.Scope
                && candidate.ParentId == item.ParentId)
            .ToList();
        var siblingIndex = siblings.FindIndex(candidate => candidate.Id == item.Id);
        var previousText = siblingIndex > 0 ? siblings[siblingIndex - 1].Text : null;
        var nextText = siblingIndex >= 0 && siblingIndex + 1 < siblings.Count
            ? siblings[siblingIndex + 1].Text
            : null;

        IReadOnlyList<string>? sameRowTexts = null;
        IReadOnlyList<TemplateMigrationSemanticObservation>? selectableChildren = null;
        if (item.Kind == "table-cell" && item.Topology is not null)
        {
            sameRowTexts = objects
                .Where(candidate => candidate.Kind == "table-cell"
                    && candidate.Id != item.Id
                    && candidate.Topology?.ContainerObjectId == item.Topology.ContainerObjectId
                    && candidate.Topology.Row == item.Topology.Row
                    && !string.IsNullOrWhiteSpace(candidate.Text))
                .OrderBy(candidate => candidate.Topology!.Column)
                .Select(candidate => candidate.Text!)
                .ToList();
        }
        if (item.Kind is "paragraph" or "table-cell") selectableChildren = objects
            .Where(candidate => candidate.Kind == "run"
                && candidate.ParentId == item.Id
                && candidate.Selector is not null
                && !claimed.Contains(candidate.Id))
            .OrderBy(candidate => candidate.Id, StringComparer.Ordinal)
            .Select(ObserveForSemanticDecision)
            .ToList();

        var context = string.IsNullOrWhiteSpace(previousText)
            && string.IsNullOrWhiteSpace(nextText)
            && (sameRowTexts is null || sameRowTexts.Count == 0)
            && (selectableChildren is null || selectableChildren.Count == 0)
                ? null
                : new TemplateMigrationSemanticContext(
                    string.IsNullOrWhiteSpace(previousText) ? null : previousText,
                    string.IsNullOrWhiteSpace(nextText) ? null : nextText,
                    sameRowTexts is { Count: > 0 } ? sameRowTexts : null,
                    selectableChildren is { Count: > 0 } ? selectableChildren : null);
        return new TemplateMigrationSemanticObservation(item.Kind, item.Scope, item.Text, item.Selector, context);
    }

    private static IReadOnlyList<(TemplateMigrationObject Source, TemplateMigrationObject Baseline)> FindEqualAnchorGapCandidates(
        IReadOnlyList<TemplateMigrationObject> source,
        IReadOnlyList<TemplateMigrationObject> baseline,
        IDictionary<string, TemplateMigrationMapping> mappings,
        IReadOnlySet<string> pendingSourceIds)
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
                .Where(item => pendingSourceIds.Contains(item.Id))
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
        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            Console.WriteLine("""
Purpose: Preserve selector-level semantic candidate resolution for compatible callers and independently re-read its selectors from the current source and baseline.
Consumes: One current source DOCX, one selected baseline DOCX, and a candidate containing mappings plus any applicable append, insertion, value-projection, choice-selection, or baseline-clear branches.
Produces: A hash-bound migration plan, an empty Unresolved array on pass, or typed selector and coverage failures without document mutation.
Use when: An existing integration is explicitly bound to list-template-migration-options and selector-level candidates.
Do not use for: New Agent-facing migration work, discovering source observations, submitting an empty diagnostic candidate, inventing values or targets, building operations, editing, or closing genuine local review items. New callers use the incremental template-migration decision commands.
Usage:
  tiwater-docx resolve-template-migration-semantic-candidate <source.docx> <baseline.docx> <candidate.json>

candidate.json uses the published camel-case candidate shape. For v5:
  required: schema, mappings (the array may be empty when another branch supplies content)
  optional: bodyAppends, valueProjections, bodyInsertions, choiceSelections, baselineClears
  selector: kind plus exactly one of text, sha256, or descendantText; optional scope,
            parentText, previousText, nextText, sameRowText, sameColumnText
  mapping: source, disposition, and baseline unless disposition is out-of-scope;
           optional cardinality is one, or all only for out-of-scope

Existing branch shapes:
  bodyAppends: sourceStart, sourceEnd
  bodyInsertions: sourceStart, sourceEnd, baselineBefore, baselineAfter,
                  stylePolicy (target-after-context)
  valueProjections: sourceParent, baselineParent, semantic, valueKind, extraction
  choiceSelections: sourceMember, baselineLabel (a run selector)
  baselineClears: baseline, mode (cell or row)
  v6 empty selector: textState (empty), explicit scope, and at least one of
                     parentText, previousText, or nextText

Every value above is selected from the current source/baseline inventories.
Candidate source selectors address only items reported in Unresolved by the
current automatic plan. Plan.Mappings are already complete and must not be
repeated. AvailableTargets is the current selectable baseline inventory for
semantic mappings and baseline-only cleanup; candidate discovery does not
publish a separate target recommendation.
Every RequiredDecisions source must be addressed. An omitted source is returned
as template-migration-semantic-decision-missing; it is not reported as a target
selection failure or local business ambiguity.
Unknown fields, object ids, indexes, and coordinates are rejected.

Minimal v5 example (values are observations from the current source/baseline):
{
  "schema": "tiwater.docx.template-migration-semantic-candidate/v5",
  "mappings": [
    {
      "source": {"kind":"paragraph","scope":"body","text":"<source text>"},
      "baseline": {"kind":"paragraph","scope":"body","text":"<baseline text>"},
      "disposition": "copy-text"
    },
    {
      "source": {"kind":"paragraph","scope":"header","text":"<excluded source text>"},
      "disposition": "out-of-scope"
    }
  ],
  "choiceSelections": [
    {
      "sourceMember": {"kind":"table-cell","scope":"body","text":"<selected member>"},
      "baselineLabel": {"kind":"run","scope":"body","text":"<target label>"}
    }
  ],
  "baselineClears": [
    {
      "baseline": {"kind":"table-cell","scope":"body","text":"<baseline placeholder>"},
      "mode": "cell"
    }
  ]
}

Allowed mapping dispositions: copy-text, copy-media, retain-target,
retain-target-label, out-of-scope. Successful resolution returns Pass=true,
Plan, and an empty Unresolved array; the operation builder consumes Plan. If
genuine local ambiguity remains after all determinate mappings were proposed,
close-template-migration-reviews consumes this resolution as a separate step.
""");
            return 0;
        }
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
        return ResolveSemanticCandidate(source, baseline, candidate, analysis, automatic);
    }

    private static TemplateMigrationMappingDerivation ResolveSemanticCandidate(
        string source,
        string baseline,
        TemplateMigrationSemanticCandidate candidate,
        TemplateMigrationAnalysis analysis,
        TemplateMigrationMappingDerivation automatic)
    {
        var mappings = automatic.Plan.Mappings
            .Where(mapping => !string.Equals(mapping.Disposition, "unresolved", StringComparison.Ordinal))
            .ToDictionary(mapping => mapping.SourceObjectId, StringComparer.Ordinal);
        var pending = automatic.Unresolved
            .Where(item => !string.IsNullOrWhiteSpace(item.SourceObjectId))
            .GroupBy(item => item.SourceObjectId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var failures = new List<TemplateMigrationPlanFailure>();
        var addressedSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var bodyAppends = new List<TemplateMigrationBodyAppend>();
        var bodyInsertions = new List<TemplateMigrationBodyInsertion>();
        var choiceSelections = new List<TemplateMigrationChoiceSelection>();
        var valueProjections = new List<TemplateMigrationValueProjection>();
        var baselineClears = new List<TemplateMigrationBaselineClear>();

        foreach (var proposal in candidate.Mappings ?? [])
        {
            var sourceMatches = ResolveSelector(analysis.Source.Objects, proposal.Source);
            var allTerminalMatches = string.Equals(proposal.Cardinality, "all", StringComparison.Ordinal)
                && string.Equals(proposal.Disposition, "out-of-scope", StringComparison.Ordinal);
            if (sourceMatches.Count == 0 || (!allTerminalMatches && sourceMatches.Count != 1))
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceMatches.Count == 0 ? "template-migration-semantic-source-missing" : "template-migration-semantic-source-ambiguous", Detail: proposal.Source.Kind));
                continue;
            }
            if (allTerminalMatches)
            {
                foreach (var matchedObject in sourceMatches)
                {
                    if (!pending.ContainsKey(matchedObject.Id))
                    {
                        failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-source-not-pending", matchedObject.Id));
                        continue;
                    }
                    mappings[matchedObject.Id] = new TemplateMigrationMapping(matchedObject.Id, null, "out-of-scope", "semantic-candidate-out-of-scope-all");
                    pending.Remove(matchedObject.Id);
                }
                continue;
            }
            var sourceObject = sourceMatches[0];
            var sourcePending = pending.ContainsKey(sourceObject.Id);
            var newRunMapping = sourceObject.Kind == "run" && !mappings.ContainsKey(sourceObject.Id);
            if (!sourcePending && !newRunMapping)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-source-not-pending", sourceObject.Id));
                continue;
            }
            if (sourcePending) addressedSourceIds.Add(sourceObject.Id);
            if (string.Equals(proposal.Disposition, "out-of-scope", StringComparison.Ordinal))
            {
                mappings[sourceObject.Id] = new TemplateMigrationMapping(sourceObject.Id, null, proposal.Disposition, "semantic-candidate-out-of-scope");
                pending.Remove(sourceObject.Id);
                continue;
            }
            var baselineMatches = ResolveSelector(analysis.Baseline.Objects, proposal.Baseline!);
            if (baselineMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(baselineMatches.Count == 0 ? "template-migration-semantic-baseline-missing" : "template-migration-semantic-baseline-ambiguous", Detail: proposal.Baseline!.Kind));
                continue;
            }
            var baselineObject = baselineMatches[0];
            var reason = proposal.Disposition switch
            {
                "retain-target" => "semantic-candidate-retain-target",
                "retain-target-label" => "semantic-candidate-retain-target-label",
                _ => "semantic-candidate-resolved"
            };
            mappings[sourceObject.Id] = new TemplateMigrationMapping(sourceObject.Id, baselineObject.Id, proposal.Disposition, reason);
            pending.Remove(sourceObject.Id);
        }

        foreach (var proposal in candidate.BodyAppends ?? [])
        {
            var starts = ResolveSelector(analysis.Source.Objects, proposal.SourceStart);
            var ends = ResolveSelector(analysis.Source.Objects, proposal.SourceEnd);
            if (starts.Count != 1 || ends.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(starts.Count != 1
                    ? starts.Count == 0 ? "template-migration-semantic-append-start-missing" : "template-migration-semantic-append-start-ambiguous"
                    : ends.Count == 0 ? "template-migration-semantic-append-end-missing" : "template-migration-semantic-append-end-ambiguous"));
                continue;
            }
            var range = BodyRange(analysis.Source.Objects, starts[0].Id, ends[0].Id);
            if (range is null)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-append-range-invalid", starts[0].Id, ends[0].Id));
                continue;
            }
            var sourceObjectIds = DescendantsOf(analysis.Source.Objects, range).ToList();
            addressedSourceIds.UnionWith(sourceObjectIds.Where(pending.ContainsKey));
            foreach (var objectId in sourceObjectIds)
            {
                mappings.Remove(objectId);
                pending.Remove(objectId);
            }
            bodyAppends.Add(new TemplateMigrationBodyAppend(starts[0].Id, ends[0].Id));
        }

        foreach (var proposal in candidate.BodyInsertions ?? [])
        {
            var starts = ResolveSelector(analysis.Source.Objects, proposal.SourceStart);
            var ends = ResolveSelector(analysis.Source.Objects, proposal.SourceEnd);
            var before = ResolveSelector(analysis.Baseline.Objects, proposal.BaselineBefore);
            var after = ResolveSelector(analysis.Baseline.Objects, proposal.BaselineAfter);
            if (starts.Count != 1 || ends.Count != 1 || before.Count != 1 || after.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-body-insertion-anchor-not-unique"));
                continue;
            }
            var range = BodyRange(analysis.Source.Objects, starts[0].Id, ends[0].Id);
            var baselineRoots = analysis.Baseline.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
            var beforeIndex = baselineRoots.FindIndex(item => item.Id == before[0].Id);
            var afterIndex = baselineRoots.FindIndex(item => item.Id == after[0].Id);
            if (range is null || beforeIndex < 0 || afterIndex != beforeIndex + 1)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-body-insertion-range-invalid", starts[0].Id, after[0].Id));
                continue;
            }
            var sourceObjectIds = DescendantsOf(analysis.Source.Objects, range).ToList();
            addressedSourceIds.UnionWith(sourceObjectIds.Where(pending.ContainsKey));
            foreach (var objectId in sourceObjectIds)
            {
                mappings.Remove(objectId);
                pending.Remove(objectId);
            }
            bodyInsertions.Add(new TemplateMigrationBodyInsertion(starts[0].Id, ends[0].Id, before[0].Id, after[0].Id, proposal.StylePolicy));
        }

        foreach (var proposal in candidate.ChoiceSelections ?? [])
        {
            var members = ResolveSelector(analysis.Source.Objects, proposal.SourceMember);
            var labels = ResolveSelector(analysis.Baseline.Objects, proposal.BaselineLabel);
            if (members.Count != 1 || labels.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-choice-selector-not-unique"));
                continue;
            }
            if (pending.ContainsKey(members[0].Id)) addressedSourceIds.Add(members[0].Id);
            if (labels[0].Kind != "run" || string.IsNullOrWhiteSpace(members[0].Text))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-choice-binding-invalid", members[0].Id, labels[0].Id));
                continue;
            }
            mappings.Remove(members[0].Id);
            pending.Remove(members[0].Id);
            choiceSelections.Add(new TemplateMigrationChoiceSelection(members[0].Id, labels[0].Id));
        }

        var clearedBaselineIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var proposal in candidate.BaselineClears ?? [])
        {
            var matches = ResolveSelector(analysis.Baseline.Objects, proposal.Baseline);
            if (matches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(matches.Count == 0
                    ? "template-migration-semantic-baseline-clear-missing"
                    : "template-migration-semantic-baseline-clear-ambiguous"));
                continue;
            }
            if (!clearedBaselineIds.Add(matches[0].Id))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-baseline-clear-duplicate", BaselineObjectId: matches[0].Id));
                continue;
            }
            baselineClears.Add(new TemplateMigrationBaselineClear(matches[0].Id, proposal.Mode));
        }

        var projectedSources = new HashSet<string>(StringComparer.Ordinal);
        var projectionBindings = new HashSet<string>(StringComparer.Ordinal);
        var projectedSemantics = new HashSet<string>(StringComparer.Ordinal);
        foreach (var proposal in candidate.ValueProjections ?? [])
        {
            var sourceMatches = ResolveSelector(analysis.Source.Objects, proposal.SourceParent);
            if (sourceMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceMatches.Count == 0
                    ? "template-migration-semantic-value-source-missing"
                    : "template-migration-semantic-value-source-ambiguous"));
                continue;
            }
            var baselineMatches = ResolveSelector(analysis.Baseline.Objects, proposal.BaselineParent);
            if (baselineMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(baselineMatches.Count == 0
                    ? "template-migration-semantic-value-baseline-missing"
                    : "template-migration-semantic-value-baseline-ambiguous"));
                continue;
            }
            var sourceParent = sourceMatches[0];
            if (pending.ContainsKey(sourceParent.Id)) addressedSourceIds.Add(sourceParent.Id);
            var baselineParent = baselineMatches[0];
            if (sourceParent.Kind is not ("paragraph" or "table-cell")
                || baselineParent.Kind is not ("paragraph" or "table-cell"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-parent-kind-mismatch", sourceParent.Id, baselineParent.Id, $"{sourceParent.Kind}->{baselineParent.Kind}"));
                continue;
            }
            if (!projectionBindings.Add($"{sourceParent.Id}\u001F{baselineParent.Id}\u001F{proposal.Semantic}"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-binding-duplicate", sourceParent.Id, baselineParent.Id, proposal.Semantic));
                continue;
            }
            projectedSources.Add(sourceParent.Id);
            if (!projectedSemantics.Add(proposal.Semantic))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-identity-duplicate", sourceParent.Id, baselineParent.Id, proposal.Semantic));
                continue;
            }
            if (!TryDeriveProjectionValue(analysis.Source.Objects, sourceParent, proposal.ValueKind, proposal.Extraction, out _, out var sourceFailure))
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceFailure!, sourceParent.Id, baselineParent.Id, proposal.Semantic));
                continue;
            }
            if (!TryLocateProjectionTarget(analysis.Baseline.Objects, baselineParent, proposal.ValueKind, proposal.Extraction, out _, out var targetFailure))
            {
                failures.Add(new TemplateMigrationPlanFailure(targetFailure!, sourceParent.Id, baselineParent.Id, proposal.Semantic));
                continue;
            }
            mappings.Remove(sourceParent.Id);
            pending.Remove(sourceParent.Id);
            valueProjections.Add(new TemplateMigrationValueProjection(
                sourceParent.Id,
                baselineParent.Id,
                proposal.Semantic,
                proposal.ValueKind,
                proposal.Extraction));
        }

        var copiedMediaRelationships = DeriveCoveredDrawingRelationships(analysis, mappings.Values);
        foreach (var drawing in analysis.Source.Objects.Where(item => item.Kind == "drawing"))
        {
            if (MediaRelationshipKey(drawing, "embedRelationshipId") is { } relationshipKey && copiedMediaRelationships.Contains(relationshipKey))
            {
                mappings.Remove(drawing.Id);
                pending.Remove(drawing.Id);
            }
        }

        var combinedClearPlan = baselineClears.Count != 0
            && (choiceSelections.Count != 0 || bodyInsertions.Count != 0 || valueProjections.Count != 0 || bodyAppends.Count != 0);
        var plan = new TemplateMigrationPlan(
            combinedClearPlan ? "tiwater.docx.template-migration-plan/v7"
                : baselineClears.Count != 0 ? "tiwater.docx.template-migration-plan/v3"
                : choiceSelections.Count != 0 ? "tiwater.docx.template-migration-plan/v6"
                : bodyInsertions.Count != 0 ? "tiwater.docx.template-migration-plan/v5"
                : valueProjections.Count != 0 ? "tiwater.docx.template-migration-plan/v4"
                : bodyAppends.Count == 0 ? "tiwater.docx.template-migration-plan/v1" : "tiwater.docx.template-migration-plan/v2",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            mappings.Values.OrderBy(mapping => mapping.SourceObjectId, StringComparer.Ordinal).ToList(),
            bodyAppends,
            BaselineClears: baselineClears,
            ValueProjections: valueProjections,
            BodyInsertions: bodyInsertions,
            ChoiceSelections: choiceSelections);
        failures.AddRange(pending.Values
            .Where(item => !addressedSourceIds.Contains(item.SourceObjectId!))
            .Select(item => new TemplateMigrationPlanFailure(
                "template-migration-semantic-decision-missing",
                item.SourceObjectId,
                Detail: item.Reason,
                Source: item.Source,
                Baseline: item.Baseline,
                BaselineOptions: item.BaselineOptions)));
        if (failures.Count == 0)
        {
            var build = BuildOperations(source, baseline, plan);
            failures.AddRange(build.Failures);
        }
        return new TemplateMigrationMappingDerivation(
            "tiwater.docx.template-migration-semantic-resolution/v1",
            failures.Count == 0,
            plan,
            failures);
    }

    public static int RunCloseReviews(string[] args)
    {
        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            Console.WriteLine("""
Purpose: Close only the genuine local ambiguities left by a completed semantic-resolution attempt.
Consumes: The current source DOCX, selected baseline DOCX, that semantic-resolution receipt, and a v5 review candidate containing only source selectors with disposition review-required.
Produces: A closed non-pass plan whose remaining Unresolved entries correspond exactly to explicit review-required mappings.
Use when: All determinate mappings have already been proposed and the resolution receipt leaves only business-ambiguous current items.
Do not use for: Replacing semantic resolution, converting automatic Unresolved items in bulk, naming a target, or suppressing a selector or coverage failure.
Usage: close-template-migration-reviews <source.docx> <baseline.docx> <resolution.json> <review-candidate.json>
""");
            return 0;
        }
        if (args.Length < 4)
        {
            throw new InvalidOperationException("close-template-migration-reviews requires <source.docx> <baseline.docx> <resolution.json> <review-candidate.json>");
        }
        var resolution = JsonSerializer.Deserialize<TemplateMigrationMappingDerivation>(File.ReadAllText(Path.GetFullPath(args[2])), Json.Options)
            ?? throw new InvalidOperationException("template-migration-review-resolution-invalid");
        var candidate = ReadSemanticCandidate(args[3], allowReview: true);
        var result = CloseReviews(args[0], args[1], resolution, candidate);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return IsClosedReview(result) ? 0 : 1;
    }

    public static TemplateMigrationMappingDerivation CloseReviews(
        string source,
        string baseline,
        TemplateMigrationMappingDerivation resolution,
        TemplateMigrationSemanticCandidate candidate)
    {
        ValidateReviewCandidate(candidate);
        var analysis = Analyze(source, baseline);
        var failures = new List<TemplateMigrationPlanFailure>();
        if (!string.Equals(resolution.Schema, "tiwater.docx.template-migration-semantic-resolution/v1", StringComparison.Ordinal)
            || resolution.Pass
            || !string.Equals(resolution.Plan.SourceSha256, analysis.Source.Sha256, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(resolution.Plan.BaselineSha256, analysis.Baseline.Sha256, StringComparison.OrdinalIgnoreCase)
            || resolution.Plan.Mappings.Any(mapping => string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal)))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-review-resolution-invalid"));
        }

        var pending = new Dictionary<string, TemplateMigrationPlanFailure>(StringComparer.Ordinal);
        foreach (var unresolved in resolution.Unresolved)
        {
            if (string.IsNullOrWhiteSpace(unresolved.SourceObjectId)
                || !pending.TryAdd(unresolved.SourceObjectId, unresolved)
                || resolution.Plan.Mappings.Any(mapping => string.Equals(mapping.SourceObjectId, unresolved.SourceObjectId, StringComparison.Ordinal)))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-review-resolution-not-closable", unresolved.SourceObjectId));
            }
        }
        if (pending.Count == 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-review-items-required"));
        }

        var mappings = resolution.Plan.Mappings.ToDictionary(mapping => mapping.SourceObjectId, StringComparer.Ordinal);
        foreach (var proposal in candidate.Mappings)
        {
            var sourceMatches = ResolveSelector(analysis.Source.Objects, proposal.Source);
            if (sourceMatches.Count != 1)
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceMatches.Count == 0
                    ? "template-migration-review-source-missing"
                    : "template-migration-review-source-ambiguous"));
                continue;
            }
            var sourceObject = sourceMatches[0];
            if (!pending.Remove(sourceObject.Id, out var unresolved))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-review-source-not-unresolved", sourceObject.Id));
                continue;
            }
            mappings[sourceObject.Id] = new TemplateMigrationMapping(
                sourceObject.Id,
                null,
                "review-required",
                unresolved.Reason);
        }
        failures.AddRange(pending.Values);

        var plan = resolution.Plan with
        {
            Mappings = mappings.Values.OrderBy(mapping => mapping.SourceObjectId, StringComparer.Ordinal).ToList()
        };
        if (failures.Count == 0)
        {
            failures.AddRange(BuildOperations(source, baseline, plan).Failures);
        }
        var reviewTerminals = plan.Mappings
            .Where(mapping => string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal))
            .Select(mapping => new TemplateMigrationPlanFailure(
                mapping.Reason ?? "template-migration-review-required",
                mapping.SourceObjectId))
            .ToList();
        return new TemplateMigrationMappingDerivation(
            "tiwater.docx.template-migration-review-closure/v1",
            false,
            plan,
            failures.Concat(reviewTerminals).ToList());
    }

    private static bool IsClosedReview(TemplateMigrationMappingDerivation result)
    {
        var reviewSourceIds = result.Plan.Mappings
            .Where(mapping => string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal))
            .Select(mapping => mapping.SourceObjectId)
            .ToHashSet(StringComparer.Ordinal);
        var unresolvedSourceIds = result.Unresolved
            .Where(item => item.SourceObjectId is not null)
            .Select(item => item.SourceObjectId!)
            .ToHashSet(StringComparer.Ordinal);
        return string.Equals(result.Schema, "tiwater.docx.template-migration-review-closure/v1", StringComparison.Ordinal)
            && reviewSourceIds.Count != 0
            && result.Unresolved.Count == reviewSourceIds.Count
            && unresolvedSourceIds.SetEquals(reviewSourceIds);
    }

    private static void ValidateReviewCandidate(TemplateMigrationSemanticCandidate candidate)
    {
        if (!string.Equals(candidate.Schema, "tiwater.docx.template-migration-semantic-candidate/v5", StringComparison.Ordinal)
            || candidate.Mappings.Count == 0
            || (candidate.BodyAppends?.Count ?? 0) != 0
            || (candidate.ValueProjections?.Count ?? 0) != 0
            || (candidate.BodyInsertions?.Count ?? 0) != 0
            || (candidate.ChoiceSelections?.Count ?? 0) != 0
            || (candidate.BaselineClears?.Count ?? 0) != 0)
        {
            throw new InvalidOperationException("template-migration-review-candidate-shape-invalid");
        }
        foreach (var mapping in candidate.Mappings)
        {
            ValidateSemanticSelector(mapping.Source, "review-source", candidate.Schema);
            if (!string.Equals(mapping.Disposition, "review-required", StringComparison.Ordinal)
                || mapping.Baseline is not null
                || mapping.Cardinality is not null and not "one")
            {
                throw new InvalidOperationException("template-migration-review-candidate-mapping-invalid");
            }
        }
    }

    private static TemplateMigrationSemanticCandidate ReadSemanticCandidate(string file, bool allowReview = false)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(file)));
        ValidateSemanticCandidateJson(document.RootElement, allowReview);
        return JsonSerializer.Deserialize<TemplateMigrationSemanticCandidate>(document.RootElement.GetRawText(), Json.Options)
            ?? throw new InvalidOperationException("template-migration-semantic-candidate-invalid");
    }

    private static TemplateMigrationChoiceCandidate ReadChoiceCandidate(string file)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(file)));
        ValidateChoiceCandidateJson(document.RootElement);
        return JsonSerializer.Deserialize<TemplateMigrationChoiceCandidate>(document.RootElement.GetRawText(), Json.CamelCaseOptions)
            ?? throw new InvalidOperationException("template-migration-choice-candidate-invalid");
    }

    private static TemplateMigrationBusinessChoiceBatch ReadBusinessChoiceBatch(string file)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(file)));
        var root = document.RootElement;
        RequireOnlyFields(root, new HashSet<string>(["schema", "choices", "templateCleanup"], StringComparer.Ordinal), "template-migration-business-choice-batch");
        if (!root.TryGetProperty("choices", out var choices) || choices.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("template-migration-business-choices-invalid");
        foreach (var choice in choices.EnumerateArray())
            RequireOnlyFields(choice, new HashSet<string>(["sourceChoiceId", "action", "targetChoiceId", "cardinality"], StringComparer.Ordinal), "template-migration-business-choice");
        if (root.TryGetProperty("templateCleanup", out var cleanup))
        {
            if (cleanup.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-business-cleanup-invalid");
            foreach (var item in cleanup.EnumerateArray())
                RequireOnlyFields(item, new HashSet<string>(["targetChoiceId", "scope"], StringComparer.Ordinal), "template-migration-business-cleanup");
        }
        return JsonSerializer.Deserialize<TemplateMigrationBusinessChoiceBatch>(root.GetRawText(), Json.CamelCaseOptions)
            ?? throw new InvalidOperationException("template-migration-business-choice-batch-invalid");
    }

    private static void ValidateBusinessChoiceBatch(TemplateMigrationBusinessChoiceBatch batch)
    {
        if (!string.Equals(batch.Schema, "tiwater.docx.template-migration-business-choices/v1", StringComparison.Ordinal))
            throw new InvalidOperationException("template-migration-business-choice-schema-invalid");
        foreach (var choice in batch.Choices)
        {
            if (string.IsNullOrWhiteSpace(choice.SourceChoiceId)) throw new InvalidOperationException("template-migration-business-source-required");
            if (choice.Action is not ("place-content" or "keep-template-content" or "keep-template-label" or "select-template-option" or "exclude-source" or "review-source"))
                throw new InvalidOperationException("template-migration-business-action-invalid");
            if (choice.Cardinality is not (null or "one" or "all")) throw new InvalidOperationException("template-migration-business-cardinality-invalid");
            if (choice.Cardinality == "all" && choice.Action is not ("exclude-source" or "review-source"))
                throw new InvalidOperationException("template-migration-business-cardinality-all-terminal-only");
        }
        foreach (var cleanup in batch.TemplateCleanup ?? [])
        {
            if (string.IsNullOrWhiteSpace(cleanup.TargetChoiceId)) throw new InvalidOperationException("template-migration-business-cleanup-target-required");
            if (cleanup.Scope is not ("cell" or "row")) throw new InvalidOperationException("template-migration-business-cleanup-scope-invalid");
        }
    }

    private static void ValidateChoiceCandidateJson(JsonElement root)
    {
        RequireOnlyFields(root, new HashSet<string>(["schema", "mappings", "choiceSelections", "baselineClears"], StringComparer.Ordinal), "template-migration-choice-candidate");
        if (!root.TryGetProperty("mappings", out var mappings) || mappings.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("template-migration-choice-candidate-mappings-invalid");
        }
        foreach (var mapping in mappings.EnumerateArray())
        {
            RequireOnlyFields(mapping, new HashSet<string>(["sourceChoiceId", "targetChoiceId", "disposition", "cardinality"], StringComparer.Ordinal), "template-migration-choice-candidate-mapping");
        }
        if (root.TryGetProperty("choiceSelections", out var selections))
        {
            if (selections.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-choice-candidate-selections-invalid");
            foreach (var selection in selections.EnumerateArray())
            {
                RequireOnlyFields(selection, new HashSet<string>(["sourceChoiceId", "targetChoiceId"], StringComparer.Ordinal), "template-migration-choice-candidate-selection");
            }
        }
        if (root.TryGetProperty("baselineClears", out var clears))
        {
            if (clears.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-choice-candidate-clears-invalid");
            foreach (var clear in clears.EnumerateArray())
            {
                RequireOnlyFields(clear, new HashSet<string>(["targetChoiceId", "mode"], StringComparer.Ordinal), "template-migration-choice-candidate-clear");
            }
        }
    }

    private static void ValidateChoiceCandidate(TemplateMigrationChoiceCandidate candidate)
    {
        if (!string.Equals(candidate.Schema, "tiwater.docx.template-migration-choice-candidate/v1", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("template-migration-choice-candidate-schema-invalid");
        }
        if (candidate.Mappings.Count == 0
            && (candidate.ChoiceSelections?.Count ?? 0) == 0
            && (candidate.BaselineClears?.Count ?? 0) == 0)
        {
            throw new InvalidOperationException("template-migration-choice-candidate-content-required");
        }
        foreach (var mapping in candidate.Mappings)
        {
            if (string.IsNullOrWhiteSpace(mapping.SourceChoiceId)) throw new InvalidOperationException("template-migration-choice-source-required");
            if (mapping.Disposition is not ("copy-text" or "copy-media" or "retain-target" or "retain-target-label" or "out-of-scope"))
            {
                throw new InvalidOperationException("template-migration-choice-disposition-invalid");
            }
            if (mapping.Cardinality is not (null or "one" or "all")) throw new InvalidOperationException("template-migration-choice-cardinality-invalid");
            if (mapping.Cardinality == "all" && mapping.Disposition != "out-of-scope") throw new InvalidOperationException("template-migration-choice-cardinality-all-terminal-only");
            if (mapping.Disposition == "out-of-scope")
            {
                if (mapping.TargetChoiceId is not null) throw new InvalidOperationException("template-migration-choice-target-forbidden");
            }
            else if (string.IsNullOrWhiteSpace(mapping.TargetChoiceId))
            {
                throw new InvalidOperationException("template-migration-choice-target-required");
            }
        }
        foreach (var selection in candidate.ChoiceSelections ?? [])
        {
            if (string.IsNullOrWhiteSpace(selection.SourceChoiceId) || string.IsNullOrWhiteSpace(selection.TargetChoiceId))
            {
                throw new InvalidOperationException("template-migration-choice-selection-required");
            }
        }
        foreach (var clear in candidate.BaselineClears ?? [])
        {
            if (string.IsNullOrWhiteSpace(clear.TargetChoiceId)) throw new InvalidOperationException("template-migration-choice-clear-target-required");
            if (clear.Mode is not ("cell" or "row")) throw new InvalidOperationException("template-migration-choice-clear-mode-invalid");
        }
    }

    private static void ValidateSemanticCandidateJson(JsonElement root, bool allowReview)
    {
        RequireOnlyFields(root, new HashSet<string>(["schema", "mappings", "bodyAppends", "valueProjections", "bodyInsertions", "choiceSelections", "baselineClears"], StringComparer.Ordinal), "template-migration-semantic-candidate");
        if (!root.TryGetProperty("mappings", out var mappings) || mappings.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-mappings-invalid");
        foreach (var mapping in mappings.EnumerateArray())
        {
            RequireOnlyFields(mapping, new HashSet<string>(["source", "baseline", "disposition", "cardinality"], StringComparer.Ordinal), "template-migration-semantic-candidate-mapping");
            if (!mapping.TryGetProperty("source", out var source)) throw new InvalidOperationException("template-migration-semantic-candidate-source-missing");
            RequireOnlyFields(source, SemanticSelectorFields(), "template-migration-semantic-candidate-source");
            var terminalWithoutTarget = mapping.TryGetProperty("disposition", out var disposition)
                && (string.Equals(disposition.GetString(), "out-of-scope", StringComparison.Ordinal)
                    || allowReview && string.Equals(disposition.GetString(), "review-required", StringComparison.Ordinal));
            if (mapping.TryGetProperty("baseline", out var baseline))
            {
                if (terminalWithoutTarget) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-forbidden");
                RequireOnlyFields(baseline, SemanticSelectorFields(), "template-migration-semantic-candidate-baseline");
            }
            else if (!terminalWithoutTarget) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-missing");
        }
        if (root.TryGetProperty("bodyAppends", out var appends))
        {
            if (appends.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-body-appends-invalid");
            foreach (var append in appends.EnumerateArray())
            {
                RequireOnlyFields(append, new HashSet<string>(["sourceStart", "sourceEnd"], StringComparer.Ordinal), "template-migration-semantic-candidate-body-append");
                foreach (var side in new[] { "sourceStart", "sourceEnd" })
                {
                    if (!append.TryGetProperty(side, out var selector)) throw new InvalidOperationException($"template-migration-semantic-candidate-{side}-missing");
                    RequireOnlyFields(selector, SemanticSelectorFields(), $"template-migration-semantic-candidate-{side}");
                }
            }
        }
        if (root.TryGetProperty("valueProjections", out var projections))
        {
            if (projections.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-value-projections-invalid");
            foreach (var projection in projections.EnumerateArray())
            {
                RequireOnlyFields(projection, new HashSet<string>(["sourceParent", "baselineParent", "semantic", "valueKind", "extraction"], StringComparer.Ordinal), "template-migration-semantic-candidate-value-projection");
                foreach (var side in new[] { "sourceParent", "baselineParent" })
                {
                    if (!projection.TryGetProperty(side, out var selector)) throw new InvalidOperationException($"template-migration-semantic-candidate-{side}-missing");
                    RequireOnlyFields(selector, SemanticSelectorFields(), $"template-migration-semantic-candidate-{side}");
                }
            }
        }
        if (root.TryGetProperty("bodyInsertions", out var insertions))
        {
            if (insertions.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-body-insertions-invalid");
            foreach (var insertion in insertions.EnumerateArray())
            {
                RequireOnlyFields(insertion, new HashSet<string>(["sourceStart", "sourceEnd", "baselineBefore", "baselineAfter", "stylePolicy"], StringComparer.Ordinal), "template-migration-semantic-candidate-body-insertion");
                foreach (var side in new[] { "sourceStart", "sourceEnd", "baselineBefore", "baselineAfter" })
                {
                    if (!insertion.TryGetProperty(side, out var selector)) throw new InvalidOperationException($"template-migration-semantic-candidate-{side}-missing");
                    RequireOnlyFields(selector, SemanticSelectorFields(), $"template-migration-semantic-candidate-{side}");
                }
            }
        }
        if (root.TryGetProperty("choiceSelections", out var choices))
        {
            if (choices.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-choice-selections-invalid");
            foreach (var choice in choices.EnumerateArray())
            {
                RequireOnlyFields(choice, new HashSet<string>(["sourceMember", "baselineLabel"], StringComparer.Ordinal), "template-migration-semantic-candidate-choice-selection");
                foreach (var side in new[] { "sourceMember", "baselineLabel" })
                {
                    if (!choice.TryGetProperty(side, out var selector)) throw new InvalidOperationException($"template-migration-semantic-candidate-{side}-missing");
                    RequireOnlyFields(selector, SemanticSelectorFields(), $"template-migration-semantic-candidate-{side}");
                }
            }
        }
        if (root.TryGetProperty("baselineClears", out var clears))
        {
            if (clears.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-clears-invalid");
            foreach (var clear in clears.EnumerateArray())
            {
                RequireOnlyFields(clear, new HashSet<string>(["baseline", "mode"], StringComparer.Ordinal), "template-migration-semantic-candidate-baseline-clear");
                if (!clear.TryGetProperty("baseline", out var selector)) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-clear-selector-missing");
                RequireOnlyFields(selector, SemanticSelectorFields(), "template-migration-semantic-candidate-baseline-clear-selector");
            }
        }
    }

    private static IReadOnlySet<string> SemanticSelectorFields()
        => new HashSet<string>(
            ["kind", "scope", "text", "sha256", "parentText", "previousText", "nextText", "descendantText", "textState", "sameRowText", "sameColumnText"],
            StringComparer.Ordinal);

    private static void RequireOnlyFields(JsonElement element, IReadOnlySet<string> allowed, string label)
    {
        if (element.ValueKind != JsonValueKind.Object) throw new InvalidOperationException($"{label}-object-invalid");
        foreach (var property in element.EnumerateObject()) if (!allowed.Contains(property.Name)) throw new InvalidOperationException($"{label}-unknown-field:{property.Name}");
    }

    private static void ValidateSemanticCandidate(TemplateMigrationSemanticCandidate candidate)
    {
        if (candidate.Schema is not ("tiwater.docx.template-migration-semantic-candidate/v1" or "tiwater.docx.template-migration-semantic-candidate/v2" or "tiwater.docx.template-migration-semantic-candidate/v3" or "tiwater.docx.template-migration-semantic-candidate/v4" or "tiwater.docx.template-migration-semantic-candidate/v5" or "tiwater.docx.template-migration-semantic-candidate/v6")) throw new InvalidOperationException("template-migration-semantic-candidate-schema-invalid");
        if (string.Equals(candidate.Schema, "tiwater.docx.template-migration-semantic-candidate/v1", StringComparison.Ordinal)
            && (candidate.ValueProjections?.Count ?? 0) != 0) throw new InvalidOperationException("template-migration-semantic-candidate-v1-value-projection-forbidden");
        if (candidate.Schema is not "tiwater.docx.template-migration-semantic-candidate/v3"
            && candidate.Schema is not "tiwater.docx.template-migration-semantic-candidate/v4"
            && candidate.Schema is not "tiwater.docx.template-migration-semantic-candidate/v5"
            && candidate.Schema is not "tiwater.docx.template-migration-semantic-candidate/v6"
            && (candidate.BodyInsertions?.Count ?? 0) != 0) throw new InvalidOperationException("template-migration-semantic-candidate-body-insertion-schema-invalid");
        if (candidate.Schema is not ("tiwater.docx.template-migration-semantic-candidate/v4" or "tiwater.docx.template-migration-semantic-candidate/v5" or "tiwater.docx.template-migration-semantic-candidate/v6")
            && (candidate.ChoiceSelections?.Count ?? 0) != 0) throw new InvalidOperationException("template-migration-semantic-candidate-choice-selection-schema-invalid");
        if (candidate.Schema is not ("tiwater.docx.template-migration-semantic-candidate/v5" or "tiwater.docx.template-migration-semantic-candidate/v6")
            && (candidate.BaselineClears?.Count ?? 0) != 0) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-clear-schema-invalid");
        if ((candidate.Mappings is null || candidate.Mappings.Count == 0)
            && (candidate.BodyAppends is null || candidate.BodyAppends.Count == 0)
            && (candidate.ValueProjections is null || candidate.ValueProjections.Count == 0)
            && (candidate.BodyInsertions is null || candidate.BodyInsertions.Count == 0)
            && (candidate.ChoiceSelections is null || candidate.ChoiceSelections.Count == 0)
            && (candidate.BaselineClears is null || candidate.BaselineClears.Count == 0)) throw new InvalidOperationException("template-migration-semantic-candidate-content-required");
        foreach (var mapping in candidate.Mappings ?? [])
        {
            ValidateSemanticSelector(mapping.Source, "source", candidate.Schema);
            if (mapping.Disposition is not ("copy-text" or "copy-media" or "retain-target" or "retain-target-label" or "out-of-scope")) throw new InvalidOperationException("template-migration-semantic-candidate-disposition-invalid");
            if (mapping.Cardinality is not (null or "one" or "all")) throw new InvalidOperationException("template-migration-semantic-candidate-cardinality-invalid");
            if (string.Equals(mapping.Cardinality, "all", StringComparison.Ordinal)
                && !string.Equals(mapping.Disposition, "out-of-scope", StringComparison.Ordinal)) throw new InvalidOperationException("template-migration-semantic-candidate-cardinality-all-terminal-only");
            if (string.Equals(mapping.Disposition, "out-of-scope", StringComparison.Ordinal))
            {
                if (mapping.Baseline is not null) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-forbidden");
            }
            else
            {
                if (mapping.Baseline is null) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-missing");
                ValidateSemanticSelector(mapping.Baseline, "baseline", candidate.Schema);
            }
        }
        foreach (var append in candidate.BodyAppends ?? [])
        {
            ValidateSemanticSelector(append.SourceStart, "source-start", candidate.Schema);
            ValidateSemanticSelector(append.SourceEnd, "source-end", candidate.Schema);
        }
        foreach (var clear in candidate.BaselineClears ?? [])
        {
            ValidateSemanticSelector(clear.Baseline, "baseline-clear", candidate.Schema);
            if (clear.Mode is not ("cell" or "row")) throw new InvalidOperationException("template-migration-semantic-candidate-baseline-clear-mode-invalid");
        }
        foreach (var projection in candidate.ValueProjections ?? [])
        {
            ValidateSemanticSelector(projection.SourceParent, "value-source-parent", candidate.Schema);
            ValidateSemanticSelector(projection.BaselineParent, "value-baseline-parent", candidate.Schema);
            if (!Regex.IsMatch(projection.Semantic ?? string.Empty, "^[a-z][a-z0-9.-]{0,63}$", RegexOptions.CultureInvariant)) throw new InvalidOperationException("template-migration-semantic-value-identity-invalid");
            if (projection.ValueKind is not ("text" or "token" or "date" or "identifier" or "version")) throw new InvalidOperationException("template-migration-semantic-value-kind-invalid");
            if (projection.Extraction is not ("after-first-delimiter" or "unique-delimited-run-group" or "unique-delimited-value" or "whole-parent")) throw new InvalidOperationException("template-migration-semantic-value-extraction-invalid");
            if (string.Equals(projection.Extraction, "unique-delimited-value", StringComparison.Ordinal) && string.Equals(projection.ValueKind, "text", StringComparison.Ordinal)) throw new InvalidOperationException("template-migration-semantic-value-text-requires-parent-boundary");
        }
        foreach (var insertion in candidate.BodyInsertions ?? [])
        {
            ValidateSemanticSelector(insertion.SourceStart, "insertion-source-start", candidate.Schema);
            ValidateSemanticSelector(insertion.SourceEnd, "insertion-source-end", candidate.Schema);
            ValidateSemanticSelector(insertion.BaselineBefore, "insertion-baseline-before", candidate.Schema);
            ValidateSemanticSelector(insertion.BaselineAfter, "insertion-baseline-after", candidate.Schema);
            if (!string.Equals(insertion.StylePolicy, "target-after-context", StringComparison.Ordinal)) throw new InvalidOperationException("template-migration-semantic-body-insertion-style-policy-invalid");
        }
        foreach (var choice in candidate.ChoiceSelections ?? [])
        {
            ValidateSemanticSelector(choice.SourceMember, "choice-source-member", candidate.Schema);
            ValidateSemanticSelector(choice.BaselineLabel, "choice-baseline-label", candidate.Schema);
            if (choice.BaselineLabel.Kind != "run") throw new InvalidOperationException("template-migration-semantic-choice-label-kind-invalid");
        }
    }

    private static void ValidateSemanticSelector(
        TemplateMigrationSemanticSelector selector,
        string side,
        string schema)
    {
        if (string.IsNullOrWhiteSpace(selector.Kind)) throw new InvalidOperationException($"template-migration-semantic-{side}-kind-required");
        var text = !string.IsNullOrWhiteSpace(selector.Text);
        var sha = !string.IsNullOrWhiteSpace(selector.Sha256);
        var descendant = !string.IsNullOrWhiteSpace(selector.DescendantText);
        var textState = !string.IsNullOrWhiteSpace(selector.TextState);
        if (selector.SameRowText is not null && string.IsNullOrWhiteSpace(selector.SameRowText))
            throw new InvalidOperationException($"template-migration-semantic-{side}-same-row-text-invalid");
        if (selector.SameColumnText is not null && string.IsNullOrWhiteSpace(selector.SameColumnText))
            throw new InvalidOperationException($"template-migration-semantic-{side}-same-column-text-invalid");
        var tableContext = !string.IsNullOrWhiteSpace(selector.SameRowText) || !string.IsNullOrWhiteSpace(selector.SameColumnText);
        if (tableContext && !string.Equals(selector.Kind, "table-cell", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"template-migration-semantic-{side}-table-context-kind-invalid");
        }
        if (textState && !string.Equals(schema, "tiwater.docx.template-migration-semantic-candidate/v6", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"template-migration-semantic-{side}-text-state-schema-invalid");
        }
        if (textState && !string.Equals(selector.TextState, "empty", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"template-migration-semantic-{side}-text-state-invalid");
        }
        if ((text ? 1 : 0) + (sha ? 1 : 0) + (descendant ? 1 : 0) + (textState ? 1 : 0) != 1) throw new InvalidOperationException($"template-migration-semantic-{side}-selector-required");
        if (sha && !Regex.IsMatch(selector.Sha256!, "^[A-Fa-f0-9]{64}$", RegexOptions.CultureInvariant)) throw new InvalidOperationException($"template-migration-semantic-{side}-sha256-invalid");
        if (textState)
        {
            if (string.IsNullOrWhiteSpace(selector.Scope)) throw new InvalidOperationException($"template-migration-semantic-{side}-empty-scope-required");
            if (string.IsNullOrWhiteSpace(selector.ParentText)
                && string.IsNullOrWhiteSpace(selector.PreviousText)
                && string.IsNullOrWhiteSpace(selector.NextText))
            {
                throw new InvalidOperationException($"template-migration-semantic-{side}-empty-context-required");
            }
        }
    }

    private static List<TemplateMigrationObject> ResolveSelector(IReadOnlyList<TemplateMigrationObject> objects, TemplateMigrationSemanticSelector selector)
    {
        var normalizedText = selector.Text is null ? null : NormalizeMappingText(selector.Text);
        var normalizedParentText = selector.ParentText is null ? null : NormalizeMappingText(selector.ParentText);
        var normalizedPreviousText = selector.PreviousText is null ? null : NormalizeMappingText(selector.PreviousText);
        var normalizedNextText = selector.NextText is null ? null : NormalizeMappingText(selector.NextText);
        var normalizedDescendantText = selector.DescendantText is null ? null : NormalizeMappingText(selector.DescendantText);
        var normalizedSameRowText = selector.SameRowText is null ? null : NormalizeMappingText(selector.SameRowText);
        var normalizedSameColumnText = selector.SameColumnText is null ? null : NormalizeMappingText(selector.SameColumnText);
        var emptyText = string.Equals(selector.TextState, "empty", StringComparison.Ordinal);
        var byId = objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var siblings = objects.Where(item => string.Equals(item.Kind, selector.Kind, StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(selector.Scope) || string.Equals(item.Scope, selector.Scope, StringComparison.Ordinal)))
            .ToList();
        return siblings.Where(item =>
                (normalizedText is null || string.Equals(NormalizeMappingText(item.Text), normalizedText, StringComparison.Ordinal))
                && (!emptyText || string.IsNullOrEmpty(NormalizeMappingText(item.Text)))
                && (selector.Sha256 is null || (item.Provenance.TryGetValue("sha256", out var hash) && string.Equals(hash, selector.Sha256, StringComparison.OrdinalIgnoreCase)))
                && (normalizedDescendantText is null || HasDescendantText(objects, item, normalizedDescendantText))
                && (normalizedParentText is null || (item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent) && string.Equals(NormalizeMappingText(parent.Text), normalizedParentText, StringComparison.Ordinal)))
                && ContextTextMatches(siblings, item, -1, normalizedPreviousText)
                && ContextTextMatches(siblings, item, 1, normalizedNextText)
                && TableContextTextMatches(objects, item, normalizedSameRowText, sameRow: true)
                && TableContextTextMatches(objects, item, normalizedSameColumnText, sameRow: false))
            .OrderBy(item => item.Id, StringComparer.Ordinal)
            .ToList();
    }

    private static bool HasDescendantText(IReadOnlyList<TemplateMigrationObject> objects, TemplateMigrationObject item, string expected)
    {
        var descendants = DescendantsOf(objects, [item.Id]);
        return objects.Any(child => descendants.Contains(child.Id)
            && !string.Equals(child.Id, item.Id, StringComparison.Ordinal)
            && string.Equals(NormalizeMappingText(child.Text), expected, StringComparison.Ordinal));
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

    private static bool TableContextTextMatches(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject item,
        string? expected,
        bool sameRow)
    {
        if (expected is null) return true;
        if (item.Topology is null) return false;
        return objects.Any(candidate =>
            !string.Equals(candidate.Id, item.Id, StringComparison.Ordinal)
            && string.Equals(candidate.Kind, "table-cell", StringComparison.Ordinal)
            && candidate.Topology is not null
            && string.Equals(candidate.Topology.ContainerObjectId, item.Topology.ContainerObjectId, StringComparison.Ordinal)
            && (sameRow ? candidate.Topology.Row == item.Topology.Row : candidate.Topology.Column == item.Topology.Column)
            && string.Equals(NormalizeMappingText(candidate.Text), expected, StringComparison.Ordinal));
    }

    private static IReadOnlyList<string>? BodyRange(IReadOnlyList<TemplateMigrationObject> objects, string startId, string endId)
    {
        var roots = objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var start = roots.FindIndex(item => string.Equals(item.Id, startId, StringComparison.Ordinal));
        var end = roots.FindIndex(item => string.Equals(item.Id, endId, StringComparison.Ordinal));
        if (start < 0 || end < start) return null;
        return roots.Skip(start).Take(end - start + 1).Select(item => item.Id).ToList();
    }

    private static HashSet<string> DescendantsOf(IReadOnlyList<TemplateMigrationObject> objects, IEnumerable<string> roots)
    {
        var children = objects.Where(item => item.ParentId is not null).GroupBy(item => item.ParentId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Id).ToList(), StringComparer.Ordinal);
        var included = new HashSet<string>(roots, StringComparer.Ordinal);
        var pending = new Queue<string>(included);
        while (pending.TryDequeue(out var parent))
        {
            if (!children.TryGetValue(parent, out var childIds)) continue;
            foreach (var child in childIds) if (included.Add(child)) pending.Enqueue(child);
        }
        return included;
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

        if (plan.Schema is not ("tiwater.docx.template-migration-plan/v1" or "tiwater.docx.template-migration-plan/v2" or "tiwater.docx.template-migration-plan/v3" or "tiwater.docx.template-migration-plan/v4" or "tiwater.docx.template-migration-plan/v5" or "tiwater.docx.template-migration-plan/v6" or "tiwater.docx.template-migration-plan/v7"))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-schema-invalid", Detail: plan.Schema));
        }
        if (string.Equals(plan.Schema, "tiwater.docx.template-migration-plan/v1", StringComparison.Ordinal) && (plan.BodyAppends?.Count ?? 0) != 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-v1-body-append-forbidden"));
        }
        if (plan.Schema is not ("tiwater.docx.template-migration-plan/v3" or "tiwater.docx.template-migration-plan/v7") && (plan.BaselineClears?.Count ?? 0) != 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-baseline-clear-schema-invalid"));
        }
        if (plan.Schema is not ("tiwater.docx.template-migration-plan/v4" or "tiwater.docx.template-migration-plan/v5" or "tiwater.docx.template-migration-plan/v6" or "tiwater.docx.template-migration-plan/v7") && (plan.ValueProjections?.Count ?? 0) != 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-value-projection-schema-invalid"));
        }
        if (string.Equals(plan.Schema, "tiwater.docx.template-migration-plan/v4", StringComparison.Ordinal) && (plan.ValueProjections?.Count ?? 0) == 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-v4-value-projection-required"));
        }
        if (plan.Schema is not ("tiwater.docx.template-migration-plan/v5" or "tiwater.docx.template-migration-plan/v6" or "tiwater.docx.template-migration-plan/v7") && (plan.BodyInsertions?.Count ?? 0) != 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-body-insertion-schema-invalid"));
        }
        if (string.Equals(plan.Schema, "tiwater.docx.template-migration-plan/v5", StringComparison.Ordinal) && (plan.BodyInsertions?.Count ?? 0) == 0)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-v5-body-insertion-required"));
        }
        if (plan.Schema is not ("tiwater.docx.template-migration-plan/v6" or "tiwater.docx.template-migration-plan/v7") && (plan.ChoiceSelections?.Count ?? 0) != 0) failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-choice-selection-schema-invalid"));
        if (string.Equals(plan.Schema, "tiwater.docx.template-migration-plan/v6", StringComparison.Ordinal) && (plan.ChoiceSelections?.Count ?? 0) == 0) failures.Add(new TemplateMigrationPlanFailure("template-migration-plan-v6-choice-selection-required"));
        if (!string.Equals(plan.SourceSha256, analysis.Source.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-source-hash-mismatch", Detail: plan.SourceSha256));
        }
        if (!string.Equals(plan.BaselineSha256, analysis.Baseline.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-hash-mismatch", Detail: plan.BaselineSha256));
        }

        var sourceById = analysis.Source.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var sourceCellParagraphs = TableCellCopyParagraphs(source);
        var baselineCellParagraphs = TableCellCopyParagraphs(baseline);
        var baselineById = analysis.Baseline.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var mappingsBySource = new Dictionary<string, TemplateMigrationMapping>(StringComparer.Ordinal);
        var copyTargets = new HashSet<string>(StringComparer.Ordinal);
        var clearTargets = new HashSet<string>(StringComparer.Ordinal);
        var appendedSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var insertedSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var choiceSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var projectedSourceIds = new HashSet<string>(StringComparer.Ordinal);
        var projectedTargetParents = new HashSet<string>(StringComparer.Ordinal);
        var projectedSemantics = new HashSet<string>(StringComparer.Ordinal);
        var projectionBindings = new HashSet<string>(StringComparer.Ordinal);
        var bodyAppends = new List<TemplateMigrationBodyAppend>();
        var bodyInsertions = new List<TemplateMigrationBodyInsertion>();

        foreach (var clear in plan.BaselineClears ?? [])
        {
            if (!baselineById.TryGetValue(clear.BaselineObjectId, out var selected)
                || selected.Kind is not ("paragraph" or "table-cell"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-clear-object-invalid", BaselineObjectId: clear.BaselineObjectId));
                continue;
            }
            IReadOnlyList<TemplateMigrationObject> targets;
            if (string.Equals(clear.Mode, "cell", StringComparison.Ordinal))
            {
                targets = [selected];
            }
            else if (string.Equals(clear.Mode, "row", StringComparison.Ordinal)
                && selected.Kind == "table-cell"
                && selected.Topology is not null)
            {
                targets = analysis.Baseline.Objects
                    .Where(item => item.Kind == "table-cell"
                        && item.Topology?.ContainerObjectId == selected.Topology.ContainerObjectId
                        && item.Topology.Row == selected.Topology.Row)
                    .OrderBy(item => item.Topology!.Column)
                    .ToList();
            }
            else
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-clear-mode-invalid", BaselineObjectId: clear.BaselineObjectId, Detail: clear.Mode));
                continue;
            }
            foreach (var target in targets)
            {
                if (!clearTargets.Add(target.Id))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-clear-duplicate", BaselineObjectId: target.Id));
                    continue;
                }
                var operation = BuildCopyTextOperation(target.Id, string.Empty);
                if (operation is null)
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-clear-operation-unsupported", BaselineObjectId: target.Id));
                    continue;
                }
                operations.Add(operation);
            }
        }

        foreach (var append in plan.BodyAppends ?? [])
        {
            var range = BodyRange(analysis.Source.Objects, append.SourceStartObjectId, append.SourceEndObjectId);
            if (range is null)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-range-invalid", append.SourceStartObjectId, append.SourceEndObjectId));
                continue;
            }
            var covered = DescendantsOf(analysis.Source.Objects, range);
            if (covered.Overlaps(appendedSourceIds))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-overlap", append.SourceStartObjectId, append.SourceEndObjectId));
                continue;
            }
            appendedSourceIds.UnionWith(covered);
            bodyAppends.Add(append);
        }
        failures.AddRange(ValidateBodyAppendContent(source, analysis.Source, bodyAppends));

        var baselineBodyRoots = analysis.Baseline.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        foreach (var insertion in plan.BodyInsertions ?? [])
        {
            var range = BodyRange(analysis.Source.Objects, insertion.SourceStartObjectId, insertion.SourceEndObjectId);
            var beforeIndex = baselineBodyRoots.FindIndex(item => item.Id == insertion.BaselineBeforeObjectId);
            var afterIndex = baselineBodyRoots.FindIndex(item => item.Id == insertion.BaselineAfterObjectId);
            if (range is null || beforeIndex < 0 || afterIndex != beforeIndex + 1 || !string.Equals(insertion.StylePolicy, "target-after-context", StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-range-invalid", insertion.SourceStartObjectId, insertion.BaselineAfterObjectId));
                continue;
            }
            var roots = analysis.Source.Objects.Where(item => range.Contains(item.Id)).ToList();
            if (roots.Any(item => item.Kind != "paragraph"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-kind-unsupported", insertion.SourceStartObjectId, insertion.BaselineAfterObjectId));
                continue;
            }
            var covered = DescendantsOf(analysis.Source.Objects, range);
            if (covered.Overlaps(insertedSourceIds) || covered.Overlaps(appendedSourceIds))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-overlap", insertion.SourceStartObjectId, insertion.BaselineAfterObjectId));
                continue;
            }
            insertedSourceIds.UnionWith(covered);
            bodyInsertions.Add(insertion);
        }
        failures.AddRange(ValidatePlainBodyInsertionContent(source, analysis.Source, bodyInsertions));

        var choiceLabels = new HashSet<string>(StringComparer.Ordinal);
        foreach (var choice in plan.ChoiceSelections ?? [])
        {
            if (!sourceById.TryGetValue(choice.SourceMemberObjectId, out var member) || string.IsNullOrWhiteSpace(member.Text)
                || !baselineById.TryGetValue(choice.BaselineLabelRunObjectId, out var label) || label.Kind != "run"
                || !choiceSourceIds.Add(member.Id) || !choiceLabels.Add(label.Id))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-choice-binding-invalid", choice.SourceMemberObjectId, choice.BaselineLabelRunObjectId));
                continue;
            }
            var operation = BuildChoiceSelectionOperation(label);
            if (operation is null)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-choice-target-invalid", member.Id, label.Id));
                continue;
            }
            operations.Add(operation);
        }

        foreach (var projection in plan.ValueProjections ?? [])
        {
            if (!sourceById.TryGetValue(projection.SourceParentObjectId, out var sourceParent)
                || !baselineById.TryGetValue(projection.BaselineParentObjectId, out var baselineParent))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-parent-unknown", projection.SourceParentObjectId, projection.BaselineParentObjectId, projection.Semantic));
                continue;
            }
            if (sourceParent.Kind is not ("paragraph" or "table-cell")
                || baselineParent.Kind is not ("paragraph" or "table-cell"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-parent-kind-mismatch", sourceParent.Id, baselineParent.Id, $"{sourceParent.Kind}->{baselineParent.Kind}"));
                continue;
            }
            if (!projectionBindings.Add($"{sourceParent.Id}\u001F{baselineParent.Id}\u001F{projection.Semantic}"))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-binding-duplicate", sourceParent.Id, baselineParent.Id, projection.Semantic));
                continue;
            }
            projectedSourceIds.Add(sourceParent.Id);
            projectedTargetParents.Add(baselineParent.Id);
            if (!projectedSemantics.Add(projection.Semantic))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-identity-duplicate", sourceParent.Id, baselineParent.Id, projection.Semantic));
                continue;
            }
            if (appendedSourceIds.Contains(sourceParent.Id) || insertedSourceIds.Contains(sourceParent.Id))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-source-duplicate", sourceParent.Id, baselineParent.Id));
                continue;
            }
            if (!TryDeriveProjectionValue(analysis.Source.Objects, sourceParent, projection.ValueKind, projection.Extraction, out var value, out var sourceFailure))
            {
                failures.Add(new TemplateMigrationPlanFailure(sourceFailure!, sourceParent.Id, baselineParent.Id, projection.Semantic));
                continue;
            }
            if (!TryLocateProjectionTarget(analysis.Baseline.Objects, baselineParent, projection.ValueKind, projection.Extraction, out var target, out var targetFailure))
            {
                failures.Add(new TemplateMigrationPlanFailure(targetFailure!, sourceParent.Id, baselineParent.Id, projection.Semantic));
                continue;
            }
            var replacements = BuildProjectionRunReplacements(target!, value!);
            foreach (var replacement in replacements)
            {
                if (!copyTargets.Add(replacement.Run.Id))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-duplicate", sourceParent.Id, replacement.Run.Id, projection.Semantic));
                    continue;
                }
                var operation = BuildCopyTextOperation(replacement.Run.Id, replacement.Text);
                if (operation is null)
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-operation-unsupported", sourceParent.Id, replacement.Run.Id, projection.Semantic));
                    continue;
                }
                operations.Add(operation);
            }
        }

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
            if (appendedSourceIds.Contains(mapping.SourceObjectId) || insertedSourceIds.Contains(mapping.SourceObjectId) || choiceSourceIds.Contains(mapping.SourceObjectId))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-source-duplicate", mapping.SourceObjectId));
                continue;
            }
            if (projectedSourceIds.Contains(mapping.SourceObjectId))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-source-mapping-conflict", mapping.SourceObjectId, mapping.BaselineObjectId));
                continue;
            }

            var disposition = mapping.Disposition?.Trim();
            if (mapping.BaselineObjectId is not null && projectedTargetParents.Contains(mapping.BaselineObjectId))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-semantic-value-baseline-mapping-conflict", mapping.SourceObjectId, mapping.BaselineObjectId));
                continue;
            }
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
                if (clearTargets.Contains(mapping.BaselineObjectId))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-clear-copy-conflict", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                var paragraphTexts = sourceCellParagraphs.GetValueOrDefault(sourceObject.Id);
                if (paragraphTexts is not null
                    && HeaderTableCellId.IsMatch(mapping.BaselineObjectId)
                    && baselineCellParagraphs.TryGetValue(mapping.BaselineObjectId, out var baselineParagraphTexts)
                    && VisibleParagraphSequencesEquivalent(paragraphTexts, baselineParagraphTexts))
                {
                    continue;
                }
                var operation = BuildCopyTextOperation(
                    mapping.BaselineObjectId,
                    paragraphTexts is null ? sourceObject.Text ?? string.Empty : string.Join("\n", paragraphTexts),
                    paragraphTexts);
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
            else if (string.Equals(disposition, "retain-target", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.Reason))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-terminal-reason-required", mapping.SourceObjectId, mapping.BaselineObjectId));
                }
                if (string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !baselineById.TryGetValue(mapping.BaselineObjectId, out var baselineObject))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-unknown", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (sourceObject.Kind is not ("paragraph" or "table-cell") || !string.Equals(sourceObject.Kind, baselineObject.Kind, StringComparison.Ordinal))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-retain-target-parent-required", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (!copyTargets.Add(mapping.BaselineObjectId))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-duplicate", mapping.SourceObjectId, mapping.BaselineObjectId));
                }
            }
            else if (string.Equals(disposition, "retain-target-label", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.Reason))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-terminal-reason-required", mapping.SourceObjectId, mapping.BaselineObjectId));
                }
                if (string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !baselineById.TryGetValue(mapping.BaselineObjectId, out var baselineObject))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-unknown", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (sourceObject.Kind is not ("paragraph" or "table-cell") || !string.Equals(sourceObject.Kind, baselineObject.Kind, StringComparison.Ordinal))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-retain-target-parent-required", mapping.SourceObjectId, mapping.BaselineObjectId));
                    continue;
                }
                if (!copyTargets.Add(mapping.BaselineObjectId))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-baseline-object-duplicate", mapping.SourceObjectId, mapping.BaselineObjectId));
                }
            }
            else if (string.Equals(disposition, "unresolved", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mapping.Reason))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-unresolved-reason-required", mapping.SourceObjectId));
                }
                else
                {
                    failures.Add(new TemplateMigrationPlanFailure(
                        "template-migration-semantic-resolution-required",
                        mapping.SourceObjectId,
                        Detail: mapping.Reason));
                }
            }
            else if (string.Equals(disposition, "review-required", StringComparison.Ordinal)
                || string.Equals(disposition, "out-of-scope", StringComparison.Ordinal))
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

        foreach (var mapping in mappingsBySource.Values.Where(item => string.Equals(item.Disposition, "retain-target", StringComparison.Ordinal)))
        {
            var sourceParent = sourceById[mapping.SourceObjectId];
            var hasMappedFactRun = mappingsBySource.Values.Any(child =>
                string.Equals(child.Disposition, "copy-text", StringComparison.Ordinal)
                && sourceById[child.SourceObjectId].Kind == "run"
                && string.Equals(sourceById[child.SourceObjectId].ParentId, sourceParent.Id, StringComparison.Ordinal)
                && child.BaselineObjectId is not null
                && baselineById[child.BaselineObjectId].Kind == "run"
                && string.Equals(baselineById[child.BaselineObjectId].ParentId, mapping.BaselineObjectId, StringComparison.Ordinal));
            if (!hasMappedFactRun)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-retain-target-fact-run-required", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
        }

        var copiedMediaRelationships = DeriveCoveredDrawingRelationships(analysis, mappingsBySource.Values);
        foreach (var sourceObject in sourceById.Values.Where(IsMigrationRequired))
        {
            var drawingCoveredByMedia = sourceObject.Kind == "drawing"
                && MediaRelationshipKey(sourceObject, "embedRelationshipId") is { } relationshipKey
                && copiedMediaRelationships.Contains(relationshipKey);
            if (!mappingsBySource.ContainsKey(sourceObject.Id) && !appendedSourceIds.Contains(sourceObject.Id) && !insertedSourceIds.Contains(sourceObject.Id) && !projectedSourceIds.Contains(sourceObject.Id) && !choiceSourceIds.Contains(sourceObject.Id) && !drawingCoveredByMedia)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-source-object-unmapped", sourceObject.Id, Detail: sourceObject.Kind));
            }
        }

        var pass = failures.Count == 0 && !reviewRequired;
        var canonicalOperations = pass ? operations : [];
        var canonicalMediaCopies = pass ? mediaCopies : [];
        var canonicalBodyAppends = pass ? bodyAppends : [];
        var previewAllowed = failures.Count == 0;
        var canonicalPreviewOperations = previewAllowed ? operations : [];
        var canonicalPreviewMediaCopies = previewAllowed ? mediaCopies : [];
        var canonicalPreviewBodyInsertions = previewAllowed ? bodyInsertions : [];
        return new TemplateMigrationOperationBuild(
            Schema: "tiwater.docx.template-migration-operation-build/v1",
            Pass: pass,
            ReviewRequired: reviewRequired,
            SourceSha256: analysis.Source.Sha256,
            BaselineSha256: analysis.Baseline.Sha256,
            OperationsSha256: pass ? HashCanonical(new { operations = canonicalOperations, mediaCopies = canonicalMediaCopies, bodyAppends = canonicalBodyAppends, bodyInsertions }) : null,
            Operations: canonicalOperations,
            MediaCopies: canonicalMediaCopies,
            BodyAppends: canonicalBodyAppends,
            BodyInsertions: canonicalPreviewBodyInsertions,
            PreviewOperationsSha256: previewAllowed ? HashCanonical(new { operations = canonicalPreviewOperations, mediaCopies = canonicalPreviewMediaCopies, bodyInsertions = canonicalPreviewBodyInsertions }) : null,
            PreviewOperations: canonicalPreviewOperations,
            PreviewMediaCopies: canonicalPreviewMediaCopies,
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

    public static int RunValidateOutput(string[] args)
    {
        if (args.Length < 4) throw new InvalidOperationException("validate-template-migration-output requires <source.docx> <baseline.docx> <plan.json> <output.docx>");
        var planPath = Path.GetFullPath(args[2]);
        var plan = JsonSerializer.Deserialize<TemplateMigrationPlan>(File.ReadAllText(planPath), Json.Options)
            ?? throw new InvalidOperationException("template-migration-plan-invalid");
        var result = ValidateOutput(args[0], args[1], planPath, args[3], plan);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    /// <summary>Independently rebuilds authority inventories and never consumes apply-result readback.</summary>
    public static TemplateMigrationOutputValidation ValidateOutput(string source, string baseline, string planPath, string output, TemplateMigrationPlan plan)
    {
        var sourcePath = Path.GetFullPath(source); var baselinePath = Path.GetFullPath(baseline); var outputPath = Path.GetFullPath(output); var canonicalPlanPath = Path.GetFullPath(planPath);
        var build = BuildOperations(sourcePath, baselinePath, plan);
        var readback = ValidateReadback(sourcePath, baselinePath, outputPath, plan);
        var failures = new List<TemplateMigrationPlanFailure>(); failures.AddRange(build.Failures); failures.AddRange(readback.Failures);
        return new TemplateMigrationOutputValidation(
            "tiwater.docx.template-migration-output-validation/v1",
            typeof(TemplateMigration).Assembly.GetName().Version?.ToString() ?? "unknown",
            build.Pass && readback.Pass,
            sourcePath, HashFile(sourcePath), baselinePath, HashFile(baselinePath), outputPath, HashFile(outputPath), canonicalPlanPath, HashFile(canonicalPlanPath),
            build, readback, failures);
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
        var insertionFailures = ApplyBodyInsertions(source, baseline, candidatePath, build.BodyInsertions);
        var appendFailures = ApplyBodyAppends(source, candidatePath, build.BodyAppends);
        var readback = ValidateReadback(source, baseline, candidatePath, plan);
        var pass = edit.AppliedOperations.All(operation => operation.Applied) && mediaFailures.Count == 0 && insertionFailures.Count == 0 && appendFailures.Count == 0 && readback.Pass;
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
            MediaFailures: [.. mediaFailures, .. insertionFailures, .. appendFailures],
            Readback: readback);
    }

    public static int RunPreview(string[] args)
    {
        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            Console.WriteLine("""
Purpose: Produce and independently read back a provisional DOCX from the verified subset of a closed review-required migration.
Consumes: The current source DOCX, selected baseline DOCX, and the closed review receipt returned by resolve-template-migration-decisions. Legacy review plans remain accepted for compatibility.
Produces: A preview receipt with ReviewRequired=true, OutputVerified, output path, operation build, edit receipt, and independent readback.
Use when: Semantic resolution closed every current item but one or more items have a genuine local review-required terminal.
Do not use for: An unresolved plan, a failed semantic request, full-pass delivery, inventing a missing target, or replacing the canonical delivery decision.
Usage: preview-template-migration <source.docx> <baseline.docx> <closed-review-or-plan.json> <output.docx>
""");
            return 0;
        }
        if (args.Length < 4)
        {
            throw new InvalidOperationException("preview-template-migration requires <source.docx> <baseline.docx> <plan.json> <output.docx>");
        }
        var reviewJson = File.ReadAllText(Path.GetFullPath(args[2]));
        using var reviewDocument = JsonDocument.Parse(reviewJson);
        TemplateMigrationPlan plan;
        if (reviewDocument.RootElement.TryGetProperty("Plan", out _) || reviewDocument.RootElement.TryGetProperty("plan", out _))
        {
            var closure = JsonSerializer.Deserialize<TemplateMigrationMappingDerivation>(reviewJson, Json.Options)
                ?? throw new InvalidOperationException("template-migration-review-closure-invalid");
            if (!IsClosedReview(closure)) throw new InvalidOperationException("template-migration-review-closure-invalid");
            plan = closure.Plan;
        }
        else
        {
            plan = JsonSerializer.Deserialize<TemplateMigrationPlan>(reviewJson, Json.Options)
                ?? throw new InvalidOperationException("template-migration-plan-invalid");
        }
        var result = Preview(source: args[0], baseline: args[1], plan: plan, output: args[3]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.OutputVerified ? 0 : 1;
    }

    /// <summary>
    /// Produces a review-only candidate from verified subset operations. A
    /// plan with unresolved or review-required mappings is never reported as
    /// pass, and this method is not a substitute for Apply or a platform
    /// delivery decision.
    /// </summary>
    public static TemplateMigrationPreviewResult Preview(string source, string baseline, TemplateMigrationPlan plan, string output)
    {
        var build = BuildOperations(source, baseline, plan);
        if (build.Failures.Count != 0)
        {
            return new TemplateMigrationPreviewResult(
                Schema: "tiwater.docx.template-migration-preview/v1",
                Pass: false,
                ReviewRequired: build.ReviewRequired,
                OutputVerified: false,
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
        var edit = Editor.Apply(Path.GetFullPath(baseline), candidatePath, build.PreviewOperations);
        var mediaFailures = ApplyMediaCopies(source, candidatePath, build.PreviewMediaCopies);
        var insertionFailures = ApplyBodyInsertions(source, baseline, candidatePath, build.BodyInsertions);
        var appendFailures = ApplyBodyAppends(source, candidatePath, build.BodyAppends);
        var readback = ValidateReadback(source, baseline, candidatePath, plan);
        var outputVerified = edit.AppliedOperations.All(operation => operation.Applied) && mediaFailures.Count == 0 && insertionFailures.Count == 0 && appendFailures.Count == 0 && readback.Pass;
        if (outputVerified)
        {
            File.Move(candidatePath, outputPath, true);
        }
        else if (File.Exists(candidatePath))
        {
            File.Delete(candidatePath);
        }
        return new TemplateMigrationPreviewResult(
            Schema: "tiwater.docx.template-migration-preview/v1",
            Pass: build.Pass && outputVerified,
            ReviewRequired: build.ReviewRequired,
            OutputVerified: outputVerified,
            Output: outputVerified ? outputPath : null,
            Build: build,
            Edit: edit,
            MediaFailures: [.. mediaFailures, .. insertionFailures, .. appendFailures],
            Readback: readback);
    }

    /// <summary>
    /// Rebuilds both authority inventories and validates the final document;
    /// it does not trust the builder or Editor result as proof of correctness.
    /// </summary>
    public static TemplateMigrationReadback ValidateReadback(string source, string baseline, string output, TemplateMigrationPlan plan)
    {
        var sourceInventory = Inventory(source);
        var baselineAuthorityInventory = Inventory(baseline);
        var outputAuthorityInventory = Inventory(output);
        var baselineInventory = CanonicalReadbackInventory(baseline);
        var outputInventory = CanonicalReadbackInventory(output);
        var failures = new List<TemplateMigrationPlanFailure>();

        if (!string.Equals(plan.SourceSha256, sourceInventory.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-source-hash-mismatch"));
        }
        if (!string.Equals(plan.BaselineSha256, baselineAuthorityInventory.Sha256, StringComparison.Ordinal))
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-hash-mismatch"));
        }

        var sourceById = sourceInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var outputById = outputInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var sourceVisibleCellText = ReadbackTableCellVisibleText(source);
        var outputVisibleCellText = ReadbackTableCellVisibleText(output);
        var sourceCellParagraphs = ReadbackTableCellParagraphs(source);
        var baselineCellParagraphs = ReadbackTableCellParagraphs(baseline);
        var outputCellParagraphs = ReadbackTableCellParagraphs(output);
        var baselineOutputIds = BuildBaselineOutputIdMap(sourceInventory, baselineInventory, outputInventory, plan.BodyInsertions ?? []);
        string OutputId(string baselineId) => baselineOutputIds.GetValueOrDefault(baselineId, baselineId);
        foreach (var choice in plan.ChoiceSelections ?? [])
        {
            var label = baselineInventory.Objects.SingleOrDefault(item => item.Id == choice.BaselineLabelRunObjectId);
            var member = sourceInventory.Objects.SingleOrDefault(item => item.Id == choice.SourceMemberObjectId);
            if (label is null || member is null || string.IsNullOrWhiteSpace(member.Text) || label.ParentId is null)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-choice-binding-invalid", choice.SourceMemberObjectId, choice.BaselineLabelRunObjectId));
                continue;
            }
            var siblingRuns = baselineInventory.Objects.Where(item => item.Kind == "run" && item.ParentId == label.ParentId).ToList();
            var labelIndex = siblingRuns.FindIndex(item => item.Id == label.Id);
            var outputLabel = outputById.GetValueOrDefault(OutputId(label.Id));
            if (outputLabel is null || labelIndex <= 0 || !string.Equals(outputLabel.Text, label.Text, StringComparison.Ordinal)
                || !IndependentlyChoiceSelected(output, label.Id))
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-choice-state-mismatch", choice.SourceMemberObjectId, choice.BaselineLabelRunObjectId));
        }
        var selectedChoiceMediaCount = outputInventory.Objects.Count(item => item.Kind == "media" && item.Provenance.GetValueOrDefault("sha256") == "825F8542DB7249A9BE93EFE1E9D894B3BF3A531744F3DF31F015BDC9B0AC3173");
        if (selectedChoiceMediaCount != (plan.ChoiceSelections?.Count ?? 0)) failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-choice-set-mismatch", Detail: $"expected={plan.ChoiceSelections?.Count ?? 0};actual={selectedChoiceMediaCount}"));
        foreach (var mapping in plan.Mappings ?? [])
        {
            if (!string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal) && !string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal)) continue;
            if (!sourceById.TryGetValue(mapping.SourceObjectId, out var sourceObject) || string.IsNullOrWhiteSpace(mapping.BaselineObjectId) || !outputById.TryGetValue(OutputId(mapping.BaselineObjectId), out var outputObject))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-object-missing", mapping.SourceObjectId, mapping.BaselineObjectId));
                continue;
            }
            var expectedText = sourceObject.Kind == "table-cell"
                ? sourceVisibleCellText.GetValueOrDefault(sourceObject.Id)
                : sourceObject.Text;
            var actualText = outputObject.Kind == "table-cell"
                ? outputVisibleCellText.GetValueOrDefault(OutputId(mapping.BaselineObjectId!))
                : outputObject.Text;
            if (string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal) && !string.Equals(expectedText, actualText, StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-content-mismatch", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
            if (string.Equals(mapping.Disposition, "copy-text", StringComparison.Ordinal)
                && sourceObject.Kind == "table-cell"
                && outputObject.Kind == "table-cell"
                && (!sourceCellParagraphs.TryGetValue(sourceObject.Id, out var sourceParagraphs)
                    || !baselineCellParagraphs.TryGetValue(mapping.BaselineObjectId!, out var baselineParagraphs)
                    || !outputCellParagraphs.TryGetValue(OutputId(mapping.BaselineObjectId!), out var outputParagraphs)
                    || !TableCellParagraphProjectionMatches(sourceParagraphs, baselineParagraphs, outputParagraphs)))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-table-cell-style-scaffold-drift", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
            if (string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal)
                && (!sourceObject.Provenance.TryGetValue("sha256", out var sourceHash)
                    || !outputObject.Provenance.TryGetValue("sha256", out var outputHash)
                    || !string.Equals(sourceHash, outputHash, StringComparison.Ordinal)))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-media-mismatch", mapping.SourceObjectId, mapping.BaselineObjectId));
            }
        }

        var baselineById = baselineInventory.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        foreach (var projectionGroup in (plan.ValueProjections ?? []).GroupBy(item => item.BaselineParentObjectId, StringComparer.Ordinal))
        {
            var first = projectionGroup.First();
            if (!baselineById.TryGetValue(first.BaselineParentObjectId, out var baselineParent)
                || !outputById.TryGetValue(OutputId(first.BaselineParentObjectId), out var outputParent))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-semantic-value-mismatch", first.SourceParentObjectId, first.BaselineParentObjectId, first.Semantic));
                continue;
            }
            var baselineRuns = baselineInventory.Objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, baselineParent.Id, StringComparison.Ordinal)).ToList();
            var outputRuns = outputInventory.Objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, outputParent.Id, StringComparison.Ordinal)).ToList();
            var expectedRunText = baselineRuns.ToDictionary(run => OutputId(run.Id), run => run.Text ?? string.Empty, StringComparer.Ordinal);
            var projectionFailed = false;
            foreach (var projection in projectionGroup)
            {
                if (!sourceById.TryGetValue(projection.SourceParentObjectId, out var sourceParent)
                    || !TryIndependentlyDeriveProjectionValue(sourceInventory.Objects, sourceParent.Id, projection.ValueKind, projection.Extraction, out var value)
                    || !TryIndependentlyBuildProjectionReplacements(baselineInventory.Objects, baselineParent.Id, projection.ValueKind, projection.Extraction, value, out var replacements))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-semantic-value-mismatch", projection.SourceParentObjectId, projection.BaselineParentObjectId, projection.Semantic));
                    projectionFailed = true;
                    continue;
                }
                foreach (var replacement in replacements) expectedRunText[OutputId(replacement.Key)] = replacement.Value;
            }
            if (!projectionFailed && (baselineRuns.Count != outputRuns.Count || outputRuns.Any(run => !expectedRunText.TryGetValue(run.Id, out var expected) || !string.Equals(run.Text ?? string.Empty, expected, StringComparison.Ordinal))))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-semantic-value-mismatch", first.SourceParentObjectId, first.BaselineParentObjectId, first.Semantic));
            }
            if (baselineRuns.Count != outputRuns.Count || baselineRuns.Where((run, index) =>
                    !string.Equals(OutputId(run.Id), outputRuns[index].Id, StringComparison.Ordinal)
                    || !string.Equals(run.Provenance.GetValueOrDefault("runPropertiesSha256"), outputRuns[index].Provenance.GetValueOrDefault("runPropertiesSha256"), StringComparison.Ordinal)
                    || !string.Equals(run.Provenance.GetValueOrDefault("paragraphPropertiesSha256"), outputRuns[index].Provenance.GetValueOrDefault("paragraphPropertiesSha256"), StringComparison.Ordinal))
                .Any())
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-semantic-value-style-drift", first.SourceParentObjectId, first.BaselineParentObjectId, first.Semantic));
            }
        }

        foreach (var clear in plan.BaselineClears ?? [])
        {
            var selected = baselineInventory.Objects.SingleOrDefault(item => item.Id == clear.BaselineObjectId);
            if (selected is null) continue;
            var targets = string.Equals(clear.Mode, "row", StringComparison.Ordinal) && selected.Topology is not null
                ? baselineInventory.Objects.Where(item => item.Kind == "table-cell"
                    && item.Topology?.ContainerObjectId == selected.Topology.ContainerObjectId
                    && item.Topology.Row == selected.Topology.Row)
                : [selected];
            foreach (var target in targets)
            {
                if (!outputById.TryGetValue(OutputId(target.Id), out var outputObject)
                    || !string.IsNullOrEmpty(outputObject.Text))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-clear-mismatch", BaselineObjectId: target.Id));
                }
                if (target.Kind == "table-cell"
                    && (!baselineCellParagraphs.TryGetValue(target.Id, out var baselineParagraphs)
                        || !outputCellParagraphs.TryGetValue(OutputId(target.Id), out var outputParagraphs)
                        || !TableCellParagraphProjectionMatches([new TableCellParagraphLine(string.Empty, string.Empty)], baselineParagraphs, outputParagraphs)))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-table-cell-style-scaffold-drift", BaselineObjectId: target.Id));
                }
            }
        }

        foreach (var mapping in (plan.Mappings ?? []).Where(item => item.Disposition is "retain-target" or "retain-target-label"))
        {
            if (string.IsNullOrWhiteSpace(mapping.BaselineObjectId)) continue;
            var copiedTargetRuns = (plan.Mappings ?? [])
                .Where(item => string.Equals(item.Disposition, "copy-text", StringComparison.Ordinal))
                .Select(item => item.BaselineObjectId)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToHashSet(StringComparer.Ordinal);
            foreach (var baselineRun in baselineInventory.Objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, mapping.BaselineObjectId, StringComparison.Ordinal)))
            {
                if (copiedTargetRuns.Contains(baselineRun.Id)) continue;
                if (!outputById.TryGetValue(OutputId(baselineRun.Id), out var outputRun)
                    || !string.Equals(baselineRun.Text, outputRun.Text, StringComparison.Ordinal)
                    || !string.Equals(baselineRun.Provenance.GetValueOrDefault("runPropertiesSha256"), outputRun.Provenance.GetValueOrDefault("runPropertiesSha256"), StringComparison.Ordinal))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-retained-target-run-mismatch", mapping.SourceObjectId, baselineRun.Id));
                }
            }
        }

        ValidateImmutableBaselineRuns(baselineInventory, outputInventory, plan, baselineOutputIds, failures);

        if ((plan.BodyAppends?.Count ?? 0) == 0 && (plan.BodyInsertions?.Count ?? 0) == 0)
        {
            var baselineStructure = baselineInventory.Objects
                .Where(IsStructuralObject)
                .Select(StructureFingerprint)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            var outputStructure = outputInventory.Objects
                .Where(IsStructuralObject)
                .Where(item => !(item.Kind == "media" && item.Provenance.GetValueOrDefault("sha256") == "825F8542DB7249A9BE93EFE1E9D894B3BF3A531744F3DF31F015BDC9B0AC3173"))
                .Select(StructureFingerprint)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            if (!baselineStructure.SequenceEqual(outputStructure, StringComparer.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-structure-drift"));
            }
        }
        else if ((plan.BodyInsertions?.Count ?? 0) == 0)
        {
            ValidateBodyAppendReadback(sourceInventory, baselineInventory, outputInventory, plan.BodyAppends!, failures);
        }
        else
        {
            ValidateBodyInsertionReadback(sourceInventory, baselineInventory, outputInventory, plan, baselineOutputIds, failures);
            if ((plan.BodyAppends?.Count ?? 0) != 0) failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-append-combination-unsupported"));
        }

        using (var baselineDocument = WordprocessingDocument.Open(Path.GetFullPath(baselineAuthorityInventory.File), false))
        using (var outputDocument = WordprocessingDocument.Open(Path.GetFullPath(outputAuthorityInventory.File), false))
        {
            var baselineErrors = CountOpenXmlErrors(baselineDocument);
            var outputErrors = CountOpenXmlErrors(outputDocument);
            foreach (var (fingerprint, count) in outputErrors)
            {
                var baselineCount = baselineErrors.GetValueOrDefault(fingerprint);
                if (count <= baselineCount) continue;
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-openxml-new-invalid", Detail: $"{fingerprint};output={count};baseline={baselineCount}"));
            }
        }

        return new TemplateMigrationReadback(failures.Count == 0, failures);
    }

    private static void ValidateImmutableBaselineRuns(
        TemplateMigrationInventory baseline,
        TemplateMigrationInventory output,
        TemplateMigrationPlan plan,
        IReadOnlyDictionary<string, string> baselineOutputIds,
        List<TemplateMigrationPlanFailure> failures)
    {
        var outputById = output.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var mutableIds = BuildMutableBaselineIds(baseline, plan);

        foreach (var baselineRun in baseline.Objects.Where(item => item.Kind == "run" && !mutableIds.Contains(item.Id)))
        {
            var outputId = baselineOutputIds.GetValueOrDefault(baselineRun.Id, baselineRun.Id);
            if (!outputById.TryGetValue(outputId, out var outputRun)
                || !string.Equals(baselineRun.Text, outputRun.Text, StringComparison.Ordinal)
                || !string.Equals(baselineRun.Provenance.GetValueOrDefault("runPropertiesSha256"), outputRun.Provenance.GetValueOrDefault("runPropertiesSha256"), StringComparison.Ordinal)
                || !string.Equals(baselineRun.Provenance.GetValueOrDefault("paragraphPropertiesSha256"), outputRun.Provenance.GetValueOrDefault("paragraphPropertiesSha256"), StringComparison.Ordinal)
                || !string.Equals(baselineRun.Provenance.GetValueOrDefault("runContentStructureSha256"), outputRun.Provenance.GetValueOrDefault("runContentStructureSha256"), StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-content-drift", BaselineObjectId: baselineRun.Id));
            }
        }
    }

    private static HashSet<string> BuildMutableBaselineIds(TemplateMigrationInventory baseline, TemplateMigrationPlan plan)
    {
        var baselineById = baseline.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var mutableIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var mapping in (plan.Mappings ?? []).Where(item =>
                     string.Equals(item.Disposition, "copy-text", StringComparison.Ordinal)
                     && !string.IsNullOrWhiteSpace(item.BaselineObjectId)))
        {
            if (baselineById.ContainsKey(mapping.BaselineObjectId!))
                mutableIds.UnionWith(DescendantsOf(baseline.Objects, [mapping.BaselineObjectId!]));
        }
        foreach (var projection in plan.ValueProjections ?? [])
        {
            if (baselineById.ContainsKey(projection.BaselineParentObjectId))
                mutableIds.UnionWith(DescendantsOf(baseline.Objects, [projection.BaselineParentObjectId]));
        }
        foreach (var clear in plan.BaselineClears ?? [])
        {
            var selected = baseline.Objects.SingleOrDefault(item => item.Id == clear.BaselineObjectId);
            if (selected is null) continue;
            var roots = string.Equals(clear.Mode, "row", StringComparison.Ordinal) && selected.Topology is not null
                ? baseline.Objects.Where(item => item.Kind == "table-cell"
                    && item.Topology?.ContainerObjectId == selected.Topology.ContainerObjectId
                    && item.Topology.Row == selected.Topology.Row).Select(item => item.Id)
                : [selected.Id];
            mutableIds.UnionWith(DescendantsOf(baseline.Objects, roots));
        }
        foreach (var choice in plan.ChoiceSelections ?? [])
        {
            if (!baselineById.TryGetValue(choice.BaselineLabelRunObjectId, out var label) || label.ParentId is null) continue;
            var siblings = baseline.Objects.Where(item => item.Kind == "run" && item.ParentId == label.ParentId).ToList();
            var labelIndex = siblings.FindIndex(item => item.Id == label.Id);
            if (labelIndex > 0) mutableIds.Add(siblings[labelIndex - 1].Id);
        }
        return mutableIds;
    }

    private static IReadOnlyList<TemplateMigrationPlanFailure> ValidatePlainBodyInsertionContent(
        string source,
        TemplateMigrationInventory inventory,
        IReadOnlyList<TemplateMigrationBodyInsertion> insertions)
    {
        if (insertions.Count == 0) return [];
        using var document = WordprocessingDocument.Open(Path.GetFullPath(source), false);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null) return [new TemplateMigrationPlanFailure("template-migration-body-insertion-body-missing")];
        var roots = inventory.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var elements = body.ChildElements.Where(element => element is Paragraph or Table).ToList();
        var failures = new List<TemplateMigrationPlanFailure>();
        foreach (var insertion in insertions)
        {
            var start = roots.FindIndex(item => item.Id == insertion.SourceStartObjectId);
            var end = roots.FindIndex(item => item.Id == insertion.SourceEndObjectId);
            if (start < 0 || end < start) continue;
            for (var index = start; index <= end; index += 1)
                if (elements[index] is not Paragraph paragraph || !IsPlainInsertionParagraph(paragraph))
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-content-unsupported", roots[index].Id));
        }
        return failures;
    }

    private static bool IsPlainInsertionParagraph(Paragraph paragraph)
    {
        if (paragraph.ChildElements.Any(child => child is not ParagraphProperties and not Run)) return false;
        foreach (var run in paragraph.Elements<Run>())
            if (run.ChildElements.Any(child => child is not RunProperties and not Text and not TabChar and not Break and not CarriageReturn)) return false;
        return true;
    }

    private static IReadOnlyList<TemplateMigrationPlanFailure> ApplyBodyInsertions(
        string source,
        string baseline,
        string output,
        IReadOnlyList<TemplateMigrationBodyInsertion> insertions)
    {
        if (insertions.Count == 0) return [];
        var failures = new List<TemplateMigrationPlanFailure>();
        using var sourceDocument = WordprocessingDocument.Open(Path.GetFullPath(source), false);
        using var outputDocument = WordprocessingDocument.Open(Path.GetFullPath(output), true);
        var sourceBody = sourceDocument.MainDocumentPart?.Document?.Body;
        var outputBody = outputDocument.MainDocumentPart?.Document?.Body;
        if (sourceBody is null || outputBody is null) return [new TemplateMigrationPlanFailure("template-migration-body-insertion-body-missing")];
        var sourceInventory = Inventory(source);
        var baselineInventory = Inventory(baseline);
        var sourceRoots = sourceInventory.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var baselineRoots = baselineInventory.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var sourceElements = sourceBody.ChildElements.Where(element => element is Paragraph or Table).ToList();
        var outputElements = outputBody.ChildElements.Where(element => element is Paragraph or Table).ToList();
        if (sourceRoots.Count != sourceElements.Count || baselineRoots.Count != outputElements.Count) return [new TemplateMigrationPlanFailure("template-migration-body-insertion-root-order-invalid")];

        foreach (var insertion in insertions.OrderByDescending(item => baselineRoots.FindIndex(root => root.Id == item.BaselineAfterObjectId)))
        {
            var sourceStart = sourceRoots.FindIndex(item => item.Id == insertion.SourceStartObjectId);
            var sourceEnd = sourceRoots.FindIndex(item => item.Id == insertion.SourceEndObjectId);
            var afterIndex = baselineRoots.FindIndex(item => item.Id == insertion.BaselineAfterObjectId);
            if (sourceStart < 0 || sourceEnd < sourceStart || afterIndex < 0 || outputElements[afterIndex] is not Paragraph context)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-range-invalid", insertion.SourceStartObjectId, insertion.BaselineAfterObjectId));
                continue;
            }
            var anchor = outputElements[afterIndex];
            for (var index = sourceStart; index <= sourceEnd; index += 1)
            {
                if (sourceElements[index] is not Paragraph sourceParagraph || !IsPlainInsertionParagraph(sourceParagraph))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-body-insertion-content-unsupported", sourceRoots[index].Id));
                    continue;
                }
                var clone = (Paragraph)sourceParagraph.CloneNode(true);
                clone.RemoveAllChildren<ParagraphProperties>();
                if (context.ParagraphProperties is not null) clone.PrependChild((ParagraphProperties)context.ParagraphProperties.CloneNode(true));
                outputBody.InsertBefore(clone, anchor);
            }
        }
        outputDocument.MainDocumentPart!.Document.Save();
        return failures;
    }

    private static IReadOnlyList<TemplateMigrationPlanFailure> ApplyBodyAppends(string source, string output, IReadOnlyList<TemplateMigrationBodyAppend> appends)
    {
        if (appends.Count == 0) return [];
        var failures = new List<TemplateMigrationPlanFailure>();
        using var sourceDocument = WordprocessingDocument.Open(Path.GetFullPath(source), false);
        using var outputDocument = WordprocessingDocument.Open(Path.GetFullPath(output), true);
        var sourceBody = sourceDocument.MainDocumentPart?.Document?.Body;
        var outputBody = outputDocument.MainDocumentPart?.Document?.Body;
        if (sourceBody is null || outputBody is null) return [new TemplateMigrationPlanFailure("template-migration-body-append-body-missing")];
        var sourceInventory = Inventory(source);
        var sourceRoots = sourceInventory.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var elements = sourceBody.ChildElements.Where(element => element is Paragraph or Table).ToList();
        if (sourceRoots.Count != elements.Count) return [new TemplateMigrationPlanFailure("template-migration-body-append-source-order-invalid")];
        var outputStyleIds = outputDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Select(style => style.StyleId?.Value).Where(id => !string.IsNullOrWhiteSpace(id)).ToHashSet(StringComparer.Ordinal) ?? [];

        foreach (var append in appends)
        {
            var range = BodyRange(sourceInventory.Objects, append.SourceStartObjectId, append.SourceEndObjectId);
            if (range is null) { failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-range-invalid", append.SourceStartObjectId, append.SourceEndObjectId)); continue; }
            var start = sourceRoots.FindIndex(item => item.Id == append.SourceStartObjectId);
            var end = sourceRoots.FindIndex(item => item.Id == append.SourceEndObjectId);
            for (var index = start; index <= end; index += 1)
            {
                var element = elements[index];
                if (element.Descendants<Drawing>().Any() || element.Descendants<FootnoteReference>().Any() || element.Descendants<EndnoteReference>().Any()
                    || element.Descendants<SdtElement>().Any() || element.Descendants().Any(item => item.LocalName is "ins" or "del" or "moveFrom" or "moveTo")
                    || (element is Paragraph paragraph && paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-unsupported-content", sourceRoots[index].Id));
                    continue;
                }
                var requiredStyles = element.Descendants<Paragraph>().Select(item => item.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                    .Concat(element.Descendants<Run>().Select(item => item.RunProperties?.RunStyle?.Val?.Value))
                    .Where(id => !string.IsNullOrWhiteSpace(id));
                if (requiredStyles.Any(id => !outputStyleIds.Contains(id!)))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-style-missing", sourceRoots[index].Id));
                    continue;
                }
                var clone = element.CloneNode(true);
                var section = outputBody.Elements<SectionProperties>().FirstOrDefault();
                if (section is null) outputBody.AppendChild(clone);
                else outputBody.InsertBefore(clone, section);
            }
        }
        outputDocument.MainDocumentPart!.Document.Save();
        return failures;
    }

    private static IReadOnlyList<TemplateMigrationPlanFailure> ValidateBodyAppendContent(
        string source,
        TemplateMigrationInventory sourceInventory,
        IReadOnlyList<TemplateMigrationBodyAppend> appends)
    {
        if (appends.Count == 0) return [];
        var failures = new List<TemplateMigrationPlanFailure>();
        using var sourceDocument = WordprocessingDocument.Open(Path.GetFullPath(source), false);
        var sourceBody = sourceDocument.MainDocumentPart?.Document?.Body;
        if (sourceBody is null) return [new TemplateMigrationPlanFailure("template-migration-body-append-body-missing")];
        var sourceRoots = sourceInventory.Objects
            .Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table")
            .ToList();
        var elements = sourceBody.ChildElements.Where(element => element is Paragraph or Table).ToList();
        if (sourceRoots.Count != elements.Count)
        {
            return [new TemplateMigrationPlanFailure("template-migration-body-append-source-order-invalid")];
        }

        foreach (var append in appends)
        {
            var start = sourceRoots.FindIndex(item => item.Id == append.SourceStartObjectId);
            var end = sourceRoots.FindIndex(item => item.Id == append.SourceEndObjectId);
            if (start < 0 || end < start)
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-range-invalid", append.SourceStartObjectId, append.SourceEndObjectId));
                continue;
            }
            for (var index = start; index <= end; index += 1)
            {
                var element = elements[index];
                if (element.Descendants<Drawing>().Any() || element.Descendants<FootnoteReference>().Any() || element.Descendants<EndnoteReference>().Any()
                    || element.Descendants<SdtElement>().Any() || element.Descendants().Any(item => item.LocalName is "ins" or "del" or "moveFrom" or "moveTo")
                    || (element is Paragraph paragraph && paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-body-append-unsupported-content", sourceRoots[index].Id));
                }
            }
        }
        return failures;
    }

    private static IReadOnlyDictionary<string, string> BuildBaselineOutputIdMap(
        TemplateMigrationInventory source,
        TemplateMigrationInventory baseline,
        TemplateMigrationInventory output,
        IReadOnlyList<TemplateMigrationBodyInsertion> insertions)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (insertions.Count == 0) return result;
        var sourceRoots = source.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var baselineRoots = baseline.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var outputRoots = output.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var outputIndex = 0;
        foreach (var baselineRoot in baselineRoots)
        {
            foreach (var insertion in insertions.Where(item => item.BaselineAfterObjectId == baselineRoot.Id))
            {
                var start = sourceRoots.FindIndex(item => item.Id == insertion.SourceStartObjectId);
                var end = sourceRoots.FindIndex(item => item.Id == insertion.SourceEndObjectId);
                if (start >= 0 && end >= start) outputIndex += end - start + 1;
            }
            if (outputIndex >= outputRoots.Count) break;
            var outputRoot = outputRoots[outputIndex++];
            foreach (var baselineObject in baseline.Objects.Where(item => item.Id == baselineRoot.Id || item.Id.StartsWith(baselineRoot.Id + ":", StringComparison.Ordinal)))
            {
                result[baselineObject.Id] = outputRoot.Id + baselineObject.Id[baselineRoot.Id.Length..];
            }
        }
        return result;
    }

    private static void ValidateBodyInsertionReadback(
        TemplateMigrationInventory source,
        TemplateMigrationInventory baseline,
        TemplateMigrationInventory output,
        TemplateMigrationPlan plan,
        IReadOnlyDictionary<string, string> baselineOutputIds,
        List<TemplateMigrationPlanFailure> failures)
    {
        var insertions = plan.BodyInsertions ?? [];
        var mutableBaselineIds = BuildMutableBaselineIds(baseline, plan);
        var mutableOutputIds = mutableBaselineIds.Select(id => baselineOutputIds.GetValueOrDefault(id, id)).ToHashSet(StringComparer.Ordinal);
        var sourceRoots = source.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var baselineRoots = baseline.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var outputRoots = output.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var insertionsByAfter = insertions.GroupBy(item => item.BaselineAfterObjectId, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var expected = new List<(TemplateMigrationObject Root, bool Inserted, string? ContextStyle)>();
        foreach (var baselineRoot in baselineRoots)
        {
            if (insertionsByAfter.TryGetValue(baselineRoot.Id, out var anchored))
            {
                foreach (var insertion in anchored)
                {
                    var start = sourceRoots.FindIndex(item => item.Id == insertion.SourceStartObjectId);
                    var end = sourceRoots.FindIndex(item => item.Id == insertion.SourceEndObjectId);
                    if (start < 0 || end < start) { failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-insertion-range-missing", insertion.SourceStartObjectId, insertion.BaselineAfterObjectId)); return; }
                    expected.AddRange(sourceRoots.Skip(start).Take(end - start + 1).Select(item => (item, true, baselineRoot.Style)));
                }
            }
            expected.Add((baselineRoot, false, (string?)null));
        }
        if (expected.Count != outputRoots.Count)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-insertion-count-mismatch", Detail: $"expected={expected.Count};actual={outputRoots.Count}"));
            return;
        }
        for (var index = 0; index < expected.Count; index += 1)
        {
            var (expectedRoot, inserted, contextStyle) = expected[index];
            var actualRoot = outputRoots[index];
            if (inserted)
            {
                if (!string.Equals(ContentTreeFingerprint(source.Objects, expectedRoot.Id), ContentTreeFingerprint(output.Objects, actualRoot.Id), StringComparison.Ordinal)
                    || !string.Equals(actualRoot.Style, contextStyle, StringComparison.Ordinal))
                {
                    failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-insertion-content-mismatch", expectedRoot.Id, actualRoot.Id));
                }
            }
            else if (!string.Equals(
                         RelativeStructureTreeFingerprint(baseline.Objects, expectedRoot.Id, mutableBaselineIds),
                         RelativeStructureTreeFingerprint(output.Objects, actualRoot.Id, mutableOutputIds),
                         StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-structure-drift", expectedRoot.Id, actualRoot.Id));
            }
        }
    }

    private static string ContentTreeFingerprint(IReadOnlyList<TemplateMigrationObject> objects, string rootId)
    {
        var descendants = DescendantsOf(objects, [rootId]);
        var rows = objects.Where(item => descendants.Contains(item.Id)).Select(item => new
        {
            item.Kind,
            item.Text,
            RunPropertiesSha256 = item.Kind == "run" ? item.Provenance.GetValueOrDefault("runPropertiesSha256") : null
        });
        return HashCanonical(rows);
    }

    private static string RelativeStructureTreeFingerprint(IReadOnlyList<TemplateMigrationObject> objects, string rootId, IReadOnlySet<string> excludedIds)
    {
        var descendants = DescendantsOf(objects, [rootId]);
        var rows = objects.Where(item => descendants.Contains(item.Id) && !excludedIds.Contains(item.Id)).Select(item => new
        {
            RelativeId = item.Id == rootId ? string.Empty
                : item.Id.StartsWith(rootId + ":", StringComparison.Ordinal) ? item.Id[rootId.Length..]
                : "external:" + item.Id,
            item.Kind,
            item.Style,
            RelativeParent = item.ParentId is null ? null : item.ParentId == rootId ? string.Empty
                : item.ParentId.StartsWith(rootId + ":", StringComparison.Ordinal) ? item.ParentId[rootId.Length..]
                : "external:" + item.ParentId,
            item.Topology
        });
        return HashCanonical(rows);
    }

    private static void ValidateBodyAppendReadback(TemplateMigrationInventory source, TemplateMigrationInventory baseline, TemplateMigrationInventory output, IReadOnlyList<TemplateMigrationBodyAppend> appends, List<TemplateMigrationPlanFailure> failures)
    {
        var outputById = output.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        foreach (var baselineObject in baseline.Objects.Where(IsStructuralObject))
        {
            if (!outputById.TryGetValue(baselineObject.Id, out var outputObject) || !string.Equals(StructureFingerprint(baselineObject), StructureFingerprint(outputObject), StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-baseline-structure-drift", baselineObject.Id));
                return;
            }
        }
        var sourceRoots = source.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var baselineRoots = baseline.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var outputRoots = output.Objects.Where(item => item.Scope == "body" && item.ParentId is null && item.Kind is "paragraph" or "table").ToList();
        var expected = new List<TemplateMigrationObject>();
        foreach (var append in appends)
        {
            var start = sourceRoots.FindIndex(item => item.Id == append.SourceStartObjectId);
            var end = sourceRoots.FindIndex(item => item.Id == append.SourceEndObjectId);
            if (start < 0 || end < start) { failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-append-range-missing", append.SourceStartObjectId, append.SourceEndObjectId)); return; }
            expected.AddRange(sourceRoots.Skip(start).Take(end - start + 1));
        }
        if (outputRoots.Count != baselineRoots.Count + expected.Count)
        {
            failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-append-count-mismatch"));
            return;
        }
        for (var index = 0; index < expected.Count; index += 1)
        {
            var actual = outputRoots[baselineRoots.Count + index];
            if (!string.Equals(ObjectTreeFingerprint(source.Objects, expected[index].Id), ObjectTreeFingerprint(output.Objects, actual.Id), StringComparison.Ordinal))
            {
                failures.Add(new TemplateMigrationPlanFailure("template-migration-readback-body-append-content-mismatch", expected[index].Id, actual.Id));
            }
        }
    }

    private static string ObjectTreeFingerprint(IReadOnlyList<TemplateMigrationObject> objects, string rootId)
    {
        var descendants = DescendantsOf(objects, [rootId]);
        var rows = objects.Where(item => descendants.Contains(item.Id)).Select(item => new
        {
            item.Kind,
            item.Scope,
            item.Text,
            item.Style,
            Provenance = item.Provenance.OrderBy(pair => pair.Key, StringComparer.Ordinal).Select(pair => new { pair.Key, pair.Value }),
        });
        return HashCanonical(rows);
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
            if (sourceObject.Provenance.TryGetValue("sha256", out var sourceHash)
                && outputObject.Provenance.TryGetValue("sha256", out var outputHash)
                && string.Equals(sourceHash, outputHash, StringComparison.OrdinalIgnoreCase))
            {
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
        AddContentControls(objects, mainPart.Document, "body", "body");
        var headerParts = mainPart.HeaderParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ToList();
        var footerParts = mainPart.FooterParts.OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ToList();
        AddHeaderFooterObjects(objects, headerParts.Select(part => part.Header), "header");
        AddHeaderFooterObjects(objects, footerParts.Select(part => part.Footer), "footer");
        foreach (var (headerPart, index) in headerParts.Select((part, index) => (part, index)))
        {
            AddContentControls(objects, headerPart.Header, $"header:{index}", "header");
        }
        foreach (var (footerPart, index) in footerParts.Select((part, index) => (part, index)))
        {
            AddContentControls(objects, footerPart.Footer, $"footer:{index}", "footer");
        }
        AddUnsupportedPart(objects, mainPart.FootnotesPart?.Footnotes, "footnotes");
        AddUnsupportedPart(objects, mainPart.EndnotesPart?.Endnotes, "endnotes");
        AddUnsupportedPart(objects, mainPart.WordprocessingCommentsPart?.Comments, "comments");
        AddDrawingObjects(objects, mainPart.Document, "mainDocument");
        AddRevisionObjects(objects, mainPart.Document, "mainDocument");
        AddMediaObjects(objects, mainPart, "mainDocument");
        foreach (var (headerPart, index) in headerParts.Select((part, index) => (part, index)))
        {
            AddDrawingObjects(objects, headerPart.Header, $"header:{index}");
            AddRevisionObjects(objects, headerPart.Header, $"header:{index}");
            AddMediaObjects(objects, headerPart, $"header:{index}");
        }
        foreach (var (footerPart, index) in footerParts.Select((part, index) => (part, index)))
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

        return new TemplateMigrationInventory(path, HashFile(path), AttachSemanticSelectors(objects));
    }

    private static IReadOnlyList<TemplateMigrationObject> AttachSemanticSelectors(IReadOnlyList<TemplateMigrationObject> objects)
        => objects.Select(item => item with { Selector = BuildSemanticSelector(objects, item) }).ToList();

    private static TemplateMigrationSemanticSelector? BuildSemanticSelector(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject item)
        => SemanticSelectorCandidates(objects, item)
            .FirstOrDefault(selector => SelectsOnly(objects, item, selector));

    private static IEnumerable<TemplateMigrationSemanticSelector> SemanticSelectorCandidates(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject item)
    {
        TemplateMigrationSemanticSelector? selector = null;
        var emptyText = string.IsNullOrWhiteSpace(item.Text);
        if (!string.IsNullOrWhiteSpace(item.Text))
        {
            selector = new TemplateMigrationSemanticSelector(item.Kind, item.Scope, Text: item.Text);
        }
        else if (item.Kind == "media"
            && item.Provenance.TryGetValue("sha256", out var sha256)
            && !string.IsNullOrWhiteSpace(sha256))
        {
            selector = new TemplateMigrationSemanticSelector(item.Kind, item.Scope, Sha256: sha256);
        }
        else if (item.Kind is "paragraph" or "table-cell" or "run")
        {
            selector = new TemplateMigrationSemanticSelector(item.Kind, item.Scope, TextState: "empty");
        }
        if (selector is null) yield break;
        if (!emptyText) yield return selector;

        if (!emptyText && item.Kind == "table-cell" && item.Topology is not null)
        {
            foreach (var context in objects
                .Where(candidate => candidate.Id != item.Id
                    && candidate.Kind == "table-cell"
                    && candidate.Topology?.ContainerObjectId == item.Topology.ContainerObjectId
                    && candidate.Topology.Row == item.Topology.Row
                    && !string.IsNullOrWhiteSpace(candidate.Text))
                .Select(candidate => candidate.Text!)
                .Distinct(StringComparer.Ordinal))
            {
                yield return selector with { SameRowText = context };
            }
            foreach (var context in objects
                .Where(candidate => candidate.Id != item.Id
                    && candidate.Kind == "table-cell"
                    && candidate.Topology?.ContainerObjectId == item.Topology.ContainerObjectId
                    && candidate.Topology.Column == item.Topology.Column
                    && !string.IsNullOrWhiteSpace(candidate.Text))
                .Select(candidate => candidate.Text!)
                .Distinct(StringComparer.Ordinal))
            {
                yield return selector with { SameColumnText = context };
            }
        }

        var byId = objects.ToDictionary(candidate => candidate.Id, StringComparer.Ordinal);
        var parentText = item.ParentId is not null
            && byId.TryGetValue(item.ParentId, out var parent)
            && !string.IsNullOrWhiteSpace(parent.Text)
                ? parent.Text
                : null;
        var siblings = objects.Where(candidate => candidate.ParentId == item.ParentId).ToList();
        var siblingIndex = siblings.FindIndex(candidate => candidate.Id == item.Id);
        var previousText = siblingIndex > 0 && !string.IsNullOrWhiteSpace(siblings[siblingIndex - 1].Text)
            ? siblings[siblingIndex - 1].Text
            : null;
        var nextText = siblingIndex >= 0
            && siblingIndex + 1 < siblings.Count
            && !string.IsNullOrWhiteSpace(siblings[siblingIndex + 1].Text)
                ? siblings[siblingIndex + 1].Text
                : null;

        foreach (var contextual in new[]
        {
            parentText is null ? null : selector with { ParentText = parentText },
            previousText is null ? null : selector with { PreviousText = previousText },
            nextText is null ? null : selector with { NextText = nextText },
            parentText is null || previousText is null ? null : selector with { ParentText = parentText, PreviousText = previousText },
            parentText is null || nextText is null ? null : selector with { ParentText = parentText, NextText = nextText },
            previousText is null || nextText is null ? null : selector with { PreviousText = previousText, NextText = nextText },
            parentText is null || previousText is null || nextText is null
                ? null
                : selector with { ParentText = parentText, PreviousText = previousText, NextText = nextText }
        })
        {
            if (contextual is not null) yield return contextual;
        }
    }

    private static bool SelectsOnly(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject expected,
        TemplateMigrationSemanticSelector selector)
    {
        var matches = ResolveSelector(objects, selector);
        return matches.Count == 1 && matches[0].Id == expected.Id;
    }

    private static TemplateMigrationInventory CanonicalReadbackInventory(string input)
    {
        var path = Path.GetFullPath(input);
        var normalizedPath = Path.Combine(Path.GetTempPath(), $"template-migration-readback-{Guid.NewGuid():N}.docx");
        try
        {
            DocxPackageNormalizer.NormalizeForReadback(path, normalizedPath);
            var normalized = Inventory(normalizedPath);
            return new TemplateMigrationInventory(path, HashFile(path), normalized.Objects);
        }
        finally
        {
            if (File.Exists(normalizedPath)) File.Delete(normalizedPath);
        }
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
                objects.Add(Object(cellId, "table-cell", scope, rowId, string.Concat(cell.Elements<Paragraph>().SelectMany(paragraph => paragraph.Descendants<Text>()).Select(text => text.Text)).Trim(), null, EmptyProvenance,
                    new TemplateMigrationTopology(tableId, rowIndex, cellIndex)));
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

    private static void AddContentControls(List<TemplateMigrationObject> objects, OpenXmlPartRootElement? root, string idPrefix, string scope)
    {
        if (root is null) return;
        foreach (var (control, index) in root.Descendants<SdtElement>().Select((control, index) => (control, index)))
        {
            objects.Add(Object($"{idPrefix}:content-control:{index}", "content-control", scope, null, control.InnerText, null,
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["propertiesSha256"] = HashXml(control)
                }));
        }
    }

    private static void AddUnsupportedPart(List<TemplateMigrationObject> objects, OpenXmlPartRootElement? root, string kind)
    {
        if (root is null) return;
        objects.Add(Object($"mainDocument:{kind}", kind, "mainDocument", null, root.InnerText, null,
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["propertiesSha256"] = HashXml(root)
            }));
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
                ["paragraphPropertiesSha256"] = HashParagraphProperties(paragraph.ParagraphProperties),
                ["runContentStructureSha256"] = HashRunContentStructure(run)
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

    private static string HashParagraphProperties(ParagraphProperties? properties)
    {
        if (properties is null) return HashText(string.Empty);
        var canonical = (ParagraphProperties)properties.CloneNode(true);
        var numbering = canonical.NumberingProperties;
        if (numbering?.NumberingLevelReference?.Val?.Value == -1
            && numbering.NumberingId?.Val?.Value == 0)
        {
            numbering.Remove();
        }
        return HashText(canonical.OuterXml);
    }

    private static string HashRunContentStructure(Run run)
    {
        var canonical = (Run)run.CloneNode(true);
        canonical.RunProperties?.Remove();
        foreach (var text in canonical.Descendants<Text>()) text.Text = string.Empty;
        return HashText(canonical.OuterXml);
    }

    private static string HashText(string text)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(text)));
    }

    private static TemplateMigrationObject Object(string id, string kind, string scope, string? parentId, string? text, string? style, IReadOnlyDictionary<string, string> provenance, TemplateMigrationTopology? topology = null)
        => new(id, kind, scope, parentId, text, style, provenance, topology);

    private static string? ParagraphStyle(Paragraph paragraph) => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    private static bool IsContentBearing(TemplateMigrationObject item)
        => (item.Kind == "paragraph" || item.Kind == "table-cell") && !string.IsNullOrWhiteSpace(item.Text);

    private static string? MediaRelationshipKey(TemplateMigrationObject item, string provenanceKey)
        => item.Provenance.TryGetValue(provenanceKey, out var relationshipId) && !string.IsNullOrWhiteSpace(relationshipId)
            ? $"{item.Scope}\u001F{relationshipId}"
            : null;

    private static IReadOnlySet<string> DeriveCoveredDrawingRelationships(
        TemplateMigrationAnalysis analysis,
        IEnumerable<TemplateMigrationMapping> mappings)
    {
        var sourceById = analysis.Source.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var baselineById = analysis.Baseline.Objects.ToDictionary(item => item.Id, StringComparer.Ordinal);
        return mappings
            .Where(mapping => string.Equals(mapping.Disposition, "copy-media", StringComparison.Ordinal)
                && mapping.BaselineObjectId is not null
                && sourceById.TryGetValue(mapping.SourceObjectId, out var sourceMedia)
                && sourceMedia.Kind == "media"
                && baselineById.TryGetValue(mapping.BaselineObjectId, out var baselineMedia)
                && baselineMedia.Kind == "media")
            .Select(mapping => (
                Source: MediaRelationshipKey(sourceById[mapping.SourceObjectId], "relationshipId"),
                Baseline: MediaRelationshipKey(baselineById[mapping.BaselineObjectId!], "relationshipId")))
            .Where(pair => pair.Source is not null && pair.Baseline is not null)
            .Where(pair => analysis.Source.Objects.Count(item => item.Kind == "drawing" && MediaRelationshipKey(item, "embedRelationshipId") == pair.Source) == 1)
            .Where(pair => analysis.Baseline.Objects.Count(item => item.Kind == "drawing" && MediaRelationshipKey(item, "embedRelationshipId") == pair.Baseline) == 1)
            .Select(pair => pair.Source!)
            .ToHashSet(StringComparer.Ordinal);
    }

    private static bool RequiresTerminalMigrationDisposition(TemplateMigrationObject item)
        => item.Kind is "revision" or "drawing" or "media" or "footnotes" or "endnotes" or "comments" or "content-control";

    private static bool IsUnsupportedObject(TemplateMigrationObject item)
        => item.Kind is "footnotes" or "endnotes" or "comments" or "content-control";

    private static bool IsMigrationRequired(TemplateMigrationObject item)
        => IsContentBearing(item) || RequiresTerminalMigrationDisposition(item);

    private static bool IsStructuralObject(TemplateMigrationObject item) => item.Kind != "run";

    private static Dictionary<string, int> CountOpenXmlErrors(WordprocessingDocument document)
    {
        var errors = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var error in new OpenXmlValidator().Validate(document))
        {
            var fingerprint = $"{error.Id}|{error.Description}";
            errors[fingerprint] = errors.GetValueOrDefault(fingerprint) + 1;
        }
        return errors;
    }

    private static string MappingKey(TemplateMigrationObject item)
        => $"{item.Kind}\u001F{NormalizeMappingText(item.Text)}";

    private static string NormalizeMappingText(string? text)
        => Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim();

    private static DocxEditOperation? BuildChoiceSelectionOperation(TemplateMigrationObject label)
    {
        var match = BodyTableCellRunId.Match(label.Id);
        if (!match.Success) return null;
        var runIndex = int.Parse(match.Groups["run"].Value);
        if (runIndex == 0) return null;
        return new DocxEditOperation(
            "setTableCellChoiceState",
            TableIndex: int.Parse(match.Groups["table"].Value),
            RowIndex: int.Parse(match.Groups["row"].Value),
            CellIndex: int.Parse(match.Groups["cell"].Value),
            ParagraphIndex: int.Parse(match.Groups["paragraph"].Value),
            RunIndex: runIndex - 1,
            Text: "selected");
    }

    private static bool IndependentlyChoiceSelected(string output, string baselineLabelRunId)
    {
        var match = BodyTableCellRunId.Match(baselineLabelRunId);
        if (!match.Success) return false;
        using var document = WordprocessingDocument.Open(Path.GetFullPath(output), false);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null) return false;
        var table = body.Elements<Table>().ElementAtOrDefault(int.Parse(match.Groups["table"].Value));
        var row = table?.Elements<TableRow>().ElementAtOrDefault(int.Parse(match.Groups["row"].Value));
        var cell = row?.Elements<TableCell>().ElementAtOrDefault(int.Parse(match.Groups["cell"].Value));
        var paragraph = cell?.Elements<Paragraph>().ElementAtOrDefault(int.Parse(match.Groups["paragraph"].Value));
        var choiceIndex = int.Parse(match.Groups["run"].Value) - 1;
        var run = choiceIndex >= 0 ? paragraph?.Elements<Run>().ElementAtOrDefault(choiceIndex) : null;
        var blip = run?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().SingleOrDefault();
        if (blip?.Embed?.Value is null || document.MainDocumentPart is null) return false;
        var part = document.MainDocumentPart.GetPartById(blip.Embed.Value);
        using var stream = part.GetStream();
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream)) == "825F8542DB7249A9BE93EFE1E9D894B3BF3A531744F3DF31F015BDC9B0AC3173";
    }

    private static DocxEditOperation? BuildCopyTextOperation(
        string baselineObjectId,
        string text,
        IReadOnlyList<string>? paragraphTexts = null)
    {
        var bodyParagraphRun = BodyParagraphRunId.Match(baselineObjectId);
        if (bodyParagraphRun.Success)
        {
            return new DocxEditOperation("replaceParagraphRunText", ParagraphIndex: int.Parse(bodyParagraphRun.Groups["paragraph"].Value), RunIndex: int.Parse(bodyParagraphRun.Groups["run"].Value), Text: text);
        }
        var headerParagraphRun = HeaderParagraphRunId.Match(baselineObjectId);
        if (headerParagraphRun.Success)
        {
            return new DocxEditOperation("replaceHeaderParagraphRunText", HeaderIndex: int.Parse(headerParagraphRun.Groups["header"].Value), ParagraphIndex: int.Parse(headerParagraphRun.Groups["paragraph"].Value), RunIndex: int.Parse(headerParagraphRun.Groups["run"].Value), Text: text);
        }
        var footerParagraphRun = FooterParagraphRunId.Match(baselineObjectId);
        if (footerParagraphRun.Success)
        {
            return new DocxEditOperation("replaceFooterParagraphRunText", FooterIndex: int.Parse(footerParagraphRun.Groups["footer"].Value), ParagraphIndex: int.Parse(footerParagraphRun.Groups["paragraph"].Value), RunIndex: int.Parse(footerParagraphRun.Groups["run"].Value), Text: text);
        }
        var bodyTableCellRun = BodyTableCellRunId.Match(baselineObjectId);
        if (bodyTableCellRun.Success)
        {
            return new DocxEditOperation("replaceTableCellRunText", TableIndex: int.Parse(bodyTableCellRun.Groups["table"].Value), RowIndex: int.Parse(bodyTableCellRun.Groups["row"].Value), CellIndex: int.Parse(bodyTableCellRun.Groups["cell"].Value), ParagraphIndex: int.Parse(bodyTableCellRun.Groups["paragraph"].Value), RunIndex: int.Parse(bodyTableCellRun.Groups["run"].Value), Text: text);
        }
        var headerTableCellRun = HeaderTableCellRunId.Match(baselineObjectId);
        if (headerTableCellRun.Success)
        {
            return new DocxEditOperation("replaceHeaderTableCellRunText", HeaderIndex: int.Parse(headerTableCellRun.Groups["header"].Value), TableIndex: int.Parse(headerTableCellRun.Groups["table"].Value), RowIndex: int.Parse(headerTableCellRun.Groups["row"].Value), CellIndex: int.Parse(headerTableCellRun.Groups["cell"].Value), ParagraphIndex: int.Parse(headerTableCellRun.Groups["paragraph"].Value), RunIndex: int.Parse(headerTableCellRun.Groups["run"].Value), Text: text);
        }
        var footerTableCellRun = FooterTableCellRunId.Match(baselineObjectId);
        if (footerTableCellRun.Success)
        {
            return new DocxEditOperation("replaceFooterTableCellRunText", FooterIndex: int.Parse(footerTableCellRun.Groups["footer"].Value), TableIndex: int.Parse(footerTableCellRun.Groups["table"].Value), RowIndex: int.Parse(footerTableCellRun.Groups["row"].Value), CellIndex: int.Parse(footerTableCellRun.Groups["cell"].Value), ParagraphIndex: int.Parse(footerTableCellRun.Groups["paragraph"].Value), RunIndex: int.Parse(footerTableCellRun.Groups["run"].Value), Text: text);
        }
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
                Text: text,
                ParagraphTexts: paragraphTexts);
        }
        var headerTableCell = HeaderTableCellId.Match(baselineObjectId);
        if (headerTableCell.Success)
        {
            return new DocxEditOperation("replaceHeaderTableCellText", HeaderIndex: int.Parse(headerTableCell.Groups["header"].Value), TableIndex: int.Parse(headerTableCell.Groups["table"].Value), RowIndex: int.Parse(headerTableCell.Groups["row"].Value), CellIndex: int.Parse(headerTableCell.Groups["cell"].Value), Text: text, ParagraphTexts: paragraphTexts);
        }
        var footerTableCell = FooterTableCellId.Match(baselineObjectId);
        if (footerTableCell.Success)
        {
            return new DocxEditOperation("replaceFooterTableCellText", FooterIndex: int.Parse(footerTableCell.Groups["footer"].Value), TableIndex: int.Parse(footerTableCell.Groups["table"].Value), RowIndex: int.Parse(footerTableCell.Groups["row"].Value), CellIndex: int.Parse(footerTableCell.Groups["cell"].Value), Text: text, ParagraphTexts: paragraphTexts);
        }
        return null;
    }

    private static bool VisibleParagraphSequencesEquivalent(
        IReadOnlyList<string> source,
        IReadOnlyList<string> baseline)
    {
        static IEnumerable<string> Visible(IReadOnlyList<string> paragraphs)
            => paragraphs.Where(text => !string.IsNullOrWhiteSpace(text));

        return Visible(source).SequenceEqual(Visible(baseline), StringComparer.Ordinal);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<string>> TableCellCopyParagraphs(string source)
    {
        var values = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
        using var document = WordprocessingDocument.Open(source, false);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("template-migration-main-document-part-missing");

        static void AddCells(IDictionary<string, IReadOnlyList<string>> target, IEnumerable<Table> tables, string prefix)
        {
            foreach (var (table, tableIndex) in tables.Select((item, index) => (item, index)))
            foreach (var (row, rowIndex) in table.Elements<TableRow>().Select((item, index) => (item, index)))
            foreach (var (cell, cellIndex) in row.Elements<TableCell>().Select((item, index) => (item, index)))
            {
                target[$"{prefix}:table:{tableIndex}:row:{rowIndex}:cell:{cellIndex}"] = cell.Elements<Paragraph>()
                    .Select(Inspector.GetParagraphText)
                    .ToList();
            }
        }

        AddCells(values, main.Document!.Body!.Elements<Table>(), "body");
        foreach (var (part, index) in main.HeaderParts.Where(part => part.Header is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            AddCells(values, part.Header!.Elements<Table>(), $"header:{index}");
        foreach (var (part, index) in main.FooterParts.Where(part => part.Footer is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            AddCells(values, part.Footer!.Elements<Table>(), $"footer:{index}");
        return values;
    }

    private static IReadOnlyDictionary<string, string> ReadbackTableCellVisibleText(string file)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        using var document = WordprocessingDocument.Open(file, false);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("template-migration-readback-main-document-part-missing");

        static void Observe(IDictionary<string, string> target, IEnumerable<Table> tables, string prefix)
        {
            foreach (var (table, tableIndex) in tables.Select((item, index) => (item, index)))
            foreach (var (row, rowIndex) in table.Elements<TableRow>().Select((item, index) => (item, index)))
            foreach (var (cell, cellIndex) in row.Elements<TableCell>().Select((item, index) => (item, index)))
            {
                var visibleParagraphs = cell.Elements<Paragraph>()
                    .Select(Inspector.GetParagraphText)
                    .Where(text => !string.IsNullOrWhiteSpace(text));
                target[$"{prefix}:table:{tableIndex}:row:{rowIndex}:cell:{cellIndex}"] = string.Join("\n", visibleParagraphs).Trim();
            }
        }

        Observe(values, main.Document!.Body!.Elements<Table>(), "body");
        foreach (var (part, index) in main.HeaderParts.Where(part => part.Header is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            Observe(values, part.Header!.Elements<Table>(), $"header:{index}");
        foreach (var (part, index) in main.FooterParts.Where(part => part.Footer is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            Observe(values, part.Footer!.Elements<Table>(), $"footer:{index}");
        return values;
    }

    private sealed record TableCellParagraphLine(
        string Text,
        string ScaffoldSha256,
        string ShellSha256 = "",
        bool HasRuns = true,
        bool HasDrawing = false);

    private static IReadOnlyDictionary<string, IReadOnlyList<TableCellParagraphLine>> ReadbackTableCellParagraphs(string file)
    {
        var values = new Dictionary<string, IReadOnlyList<TableCellParagraphLine>>(StringComparer.Ordinal);
        using var document = WordprocessingDocument.Open(file, false);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("template-migration-readback-main-document-part-missing");

        static string ScaffoldHash(Paragraph paragraph)
        {
            var clone = (Paragraph)paragraph.CloneNode(true);
            var protectedBlankText = clone.Descendants<Run>()
                .Where(run => run.RunProperties?.GetFirstChild<Underline>() is not null)
                .Select(run => string.Concat(run.Descendants<Text>().Select(value => value.Text)))
                .Where(string.IsNullOrWhiteSpace)
                .ToList();
            foreach (var value in clone.Descendants<Text>().ToList()) value.Remove();
            foreach (var lineBreak in clone.Descendants<Break>().ToList()) lineBreak.Remove();
            return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
                clone.OuterXml + "\u001E" + string.Join("\u001F", protectedBlankText))));
        }

        static string ShellHash(Paragraph paragraph)
        {
            var clone = (Paragraph)paragraph.CloneNode(true);
            clone.RemoveAllChildren<Run>();
            return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(clone.OuterXml)));
        }

        static void Observe(IDictionary<string, IReadOnlyList<TableCellParagraphLine>> target, IEnumerable<Table> tables, string prefix)
        {
            foreach (var (table, tableIndex) in tables.Select((item, index) => (item, index)))
            foreach (var (row, rowIndex) in table.Elements<TableRow>().Select((item, index) => (item, index)))
            foreach (var (cell, cellIndex) in row.Elements<TableCell>().Select((item, index) => (item, index)))
            {
                target[$"{prefix}:table:{tableIndex}:row:{rowIndex}:cell:{cellIndex}"] = cell.Elements<Paragraph>()
                    .Select(paragraph => new TableCellParagraphLine(
                        Inspector.GetParagraphText(paragraph),
                        ScaffoldHash(paragraph),
                        ShellHash(paragraph),
                        paragraph.Descendants<Run>().Any(),
                        paragraph.Descendants<Drawing>().Any()))
                    .ToList();
            }
        }

        Observe(values, main.Document!.Body!.Elements<Table>(), "body");
        foreach (var (part, index) in main.HeaderParts.Where(part => part.Header is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            Observe(values, part.Header!.Elements<Table>(), $"header:{index}");
        foreach (var (part, index) in main.FooterParts.Where(part => part.Footer is not null)
                     .OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal).Select((part, index) => (part, index)))
            Observe(values, part.Footer!.Elements<Table>(), $"footer:{index}");
        return values;
    }

    private static bool TableCellParagraphProjectionMatches(
        IReadOnlyList<TableCellParagraphLine> source,
        IReadOnlyList<TableCellParagraphLine> baseline,
        IReadOnlyList<TableCellParagraphLine> output)
    {
        if (baseline.Count == 0 || output.Count == 0) return baseline.Count == output.Count;
        static bool Visible(string value) => !string.IsNullOrWhiteSpace(value);
        static string Comparable(string value) => Visible(value) ? value : string.Empty;
        var scores = new int[source.Count + 1, baseline.Count + 1];
        for (var sourceIndex = source.Count - 1; sourceIndex >= 0; sourceIndex -= 1)
        for (var targetIndex = baseline.Count - 1; targetIndex >= 0; targetIndex -= 1)
        {
            var skipSource = scores[sourceIndex + 1, targetIndex];
            var skipTarget = scores[sourceIndex, targetIndex + 1];
            var match = Visible(source[sourceIndex].Text) == Visible(baseline[targetIndex].Text)
                ? (Visible(source[sourceIndex].Text) ? 3 : 2) + scores[sourceIndex + 1, targetIndex + 1]
                : int.MinValue;
            scores[sourceIndex, targetIndex] = Math.Max(match, Math.Max(skipSource, skipTarget));
        }

        var sourceToTarget = Enumerable.Repeat<int?>(null, source.Count).ToArray();
        var sourceCursor = 0;
        var targetCursor = 0;
        while (sourceCursor < source.Count && targetCursor < baseline.Count)
        {
            var weight = Visible(source[sourceCursor].Text) ? 3 : 2;
            if (Visible(source[sourceCursor].Text) == Visible(baseline[targetCursor].Text)
                && scores[sourceCursor, targetCursor] == weight + scores[sourceCursor + 1, targetCursor + 1])
            {
                sourceToTarget[sourceCursor] = targetCursor;
                sourceCursor += 1;
                targetCursor += 1;
            }
            else if (scores[sourceCursor + 1, targetCursor] >= scores[sourceCursor, targetCursor + 1])
            {
                sourceCursor += 1;
            }
            else
            {
                targetCursor += 1;
            }
        }
        if (!baseline.Any(item => Visible(item.Text)) && source.Any(item => Visible(item.Text)))
        {
            var promotedSource = source.ToList().FindIndex(item => Visible(item.Text));
            var promotedTarget = baseline.ToList().FindIndex(item => !item.HasDrawing);
            if (promotedTarget < 0) promotedTarget = 0;
            for (var index = 0; index < sourceToTarget.Length; index += 1)
                if (sourceToTarget[index] == promotedTarget) sourceToTarget[index] = null;
            sourceToTarget[promotedSource] = promotedTarget;
        }

        var targetToSource = sourceToTarget
            .Select((target, index) => (target, index))
            .Where(pair => pair.target is not null)
            .ToDictionary(pair => pair.target!.Value, pair => pair.index);
        var insertions = new Dictionary<int, List<int>>();
        var insertionsAfter = new Dictionary<int, List<int>>();
        var trailingInsertions = new List<int>();
        for (var index = 0; index < source.Count; index += 1)
        {
            if (sourceToTarget[index] is not null) continue;
            var nextTarget = sourceToTarget.Skip(index + 1).FirstOrDefault(candidate => candidate is not null);
            if (nextTarget is null)
            {
                var previousTarget = sourceToTarget.Take(index).LastOrDefault(candidate => candidate is not null);
                if (previousTarget is null)
                {
                    trailingInsertions.Add(index);
                }
                else
                {
                    if (!insertionsAfter.TryGetValue(previousTarget.Value, out var after))
                    {
                        after = [];
                        insertionsAfter[previousTarget.Value] = after;
                    }
                    after.Add(index);
                }
                continue;
            }
            if (!insertions.TryGetValue(nextTarget.Value, out var before))
            {
                before = [];
                insertions[nextTarget.Value] = before;
            }
            before.Add(index);
        }

        TableCellParagraphLine TemplateFor(string text)
            => baseline.LastOrDefault(item => Visible(item.Text) == Visible(text) && !item.HasDrawing)
                ?? baseline.LastOrDefault(item => !item.HasDrawing)
                ?? baseline[^1];
        var expected = new List<TableCellParagraphLine>();
        for (var baselineIndex = 0; baselineIndex < baseline.Count; baselineIndex += 1)
        {
            if (insertions.TryGetValue(baselineIndex, out var before))
                foreach (var sourceIndex in before)
                    expected.Add(TemplateFor(source[sourceIndex].Text) with { Text = source[sourceIndex].Text });
            var baselineItem = baseline[baselineIndex];
            expected.Add(targetToSource.TryGetValue(baselineIndex, out var mappedSource)
                ? baselineItem with { Text = source[mappedSource].Text }
                : Visible(baselineItem.Text) ? baselineItem with { Text = string.Empty } : baselineItem);
            if (insertionsAfter.TryGetValue(baselineIndex, out var after))
                foreach (var sourceIndex in after)
                    expected.Add(TemplateFor(source[sourceIndex].Text) with { Text = source[sourceIndex].Text });
        }
        foreach (var sourceIndex in trailingInsertions)
            expected.Add(TemplateFor(source[sourceIndex].Text) with { Text = source[sourceIndex].Text });

        return expected.Count == output.Count && expected.Zip(output).All(pair =>
            string.Equals(Comparable(pair.First.Text), Comparable(pair.Second.Text), StringComparison.Ordinal)
            && string.Equals(pair.First.ShellSha256, pair.Second.ShellSha256, StringComparison.Ordinal)
            && (!pair.First.HasRuns || string.Equals(pair.First.ScaffoldSha256, pair.Second.ScaffoldSha256, StringComparison.Ordinal)));
    }

    private sealed record ProjectionTargetSpan(
        IReadOnlyList<TemplateMigrationObject> Runs,
        int ValueStart,
        int ValueEnd);

    private sealed record ProjectionRunReplacement(TemplateMigrationObject Run, string Text);

    private static bool TryDeriveProjectionValue(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject parent,
        string valueKind,
        string extraction,
        out string? value,
        out string? failure)
    {
        value = null;
        failure = null;
        if (extraction is not ("after-first-delimiter" or "unique-delimited-run-group" or "unique-delimited-value" or "whole-parent"))
        {
            failure = "template-migration-semantic-value-extraction-invalid";
            return false;
        }
        var runs = objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, parent.Id, StringComparison.Ordinal)).ToList();
        if (runs.Count == 0)
        {
            failure = "template-migration-semantic-value-source-runs-missing";
            return false;
        }
        if (string.Equals(extraction, "whole-parent", StringComparison.Ordinal))
        {
            var observed = string.Concat(runs.Select(item => item.Text ?? string.Empty)).Trim();
            if (observed.Length == 0) { failure = "template-migration-semantic-value-source-empty"; return false; }
            if (!ProjectionValueKindMatches(observed, valueKind)) { failure = "template-migration-semantic-value-source-kind-mismatch"; return false; }
            value = observed;
            return true;
        }
        if (string.Equals(extraction, "unique-delimited-value", StringComparison.Ordinal))
        {
            var spans = ProjectionRunGroups(runs)
                .SelectMany(group => DelimitedValueSpans(string.Concat(group.Select(item => item.Text ?? string.Empty)), valueKind, allowPlaceholder: false))
                .Where(span => ProjectionValueKindMatches(span.Value, valueKind)).ToList();
            if (spans.Count == 0)
            {
                failure = "template-migration-semantic-value-source-kind-mismatch";
                return false;
            }
            if (spans.Count != 1)
            {
                failure = "template-migration-semantic-value-source-value-ambiguous";
                return false;
            }
            value = spans[0].Value;
            return true;
        }
        var groups = string.Equals(extraction, "unique-delimited-run-group", StringComparison.Ordinal)
            ? ProjectionRunGroups(runs)
            : [runs];
        var candidates = groups.Select(group =>
            {
                var text = string.Concat(group.Select(item => item.Text ?? string.Empty));
                var delimiter = FirstDelimiter(text);
                return delimiter < 0 ? null : text[(delimiter + 1)..].Trim();
            })
            .Where(candidate => !string.IsNullOrEmpty(candidate) && ProjectionValueKindMatches(candidate!, valueKind))
            .ToList();
        if (candidates.Count == 0)
        {
            var hasValue = groups.Any(group =>
            {
                var text = string.Concat(group.Select(item => item.Text ?? string.Empty));
                var delimiter = FirstDelimiter(text);
                return delimiter >= 0 && text[(delimiter + 1)..].Trim().Length != 0;
            });
            failure = hasValue ? "template-migration-semantic-value-source-kind-mismatch" : "template-migration-semantic-value-source-empty";
            return false;
        }
        if (candidates.Count != 1)
        {
            failure = "template-migration-semantic-value-source-value-ambiguous";
            return false;
        }
        value = candidates[0];
        return true;
    }

    private static bool TryLocateProjectionTarget(
        IReadOnlyList<TemplateMigrationObject> objects,
        TemplateMigrationObject parent,
        string valueKind,
        string extraction,
        out ProjectionTargetSpan? target,
        out string? failure)
    {
        target = null;
        failure = null;
        if (extraction is not ("after-first-delimiter" or "unique-delimited-run-group" or "unique-delimited-value" or "whole-parent"))
        {
            failure = "template-migration-semantic-value-extraction-invalid";
            return false;
        }
        var runs = objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, parent.Id, StringComparison.Ordinal)).ToList();
        if (runs.Count == 0)
        {
            failure = "template-migration-semantic-value-baseline-runs-missing";
            return false;
        }
        if (string.Equals(extraction, "whole-parent", StringComparison.Ordinal))
        {
            var text = string.Concat(runs.Select(item => item.Text ?? string.Empty));
            var start = 0; while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
            var end = text.Length; while (end > start && char.IsWhiteSpace(text[end - 1])) end -= 1;
            if (start == end || (!ProjectionValueKindMatches(text[start..end], valueKind) && !IsProjectionPlaceholder(text[start..end]))) { failure = "template-migration-semantic-value-baseline-empty"; return false; }
            target = new ProjectionTargetSpan(runs, start, end);
            return true;
        }
        if (string.Equals(extraction, "unique-delimited-value", StringComparison.Ordinal))
        {
            var spans = ProjectionRunGroups(runs)
                .SelectMany(group => DelimitedValueSpans(string.Concat(group.Select(item => item.Text ?? string.Empty)), valueKind, allowPlaceholder: true)
                    .Where(span => ProjectionValueKindMatches(span.Value, valueKind) || IsProjectionPlaceholder(span.Value))
                    .Select(span => (Runs: group, Span: span)))
                .ToList();
            if (spans.Count == 0)
            {
                failure = "template-migration-semantic-value-baseline-empty";
                return false;
            }
            if (spans.Count != 1)
            {
                failure = "template-migration-semantic-value-baseline-value-ambiguous";
                return false;
            }
            target = new ProjectionTargetSpan(spans[0].Runs, spans[0].Span.Start, spans[0].Span.End);
            return true;
        }
        var groups = string.Equals(extraction, "unique-delimited-run-group", StringComparison.Ordinal)
            ? ProjectionRunGroups(runs)
            : [runs];
        var candidates = new List<ProjectionTargetSpan>();
        foreach (var group in groups)
        {
            var text = string.Concat(group.Select(item => item.Text ?? string.Empty));
            var delimiter = FirstDelimiter(text);
            if (delimiter < 0) continue;
            var start = delimiter + 1;
            while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
            var end = text.Length;
            while (end > start && char.IsWhiteSpace(text[end - 1])) end -= 1;
            if (start >= end) continue;
            var observed = text[start..end];
            if (!ProjectionValueKindMatches(observed, valueKind) && !IsProjectionPlaceholder(observed)) continue;
            candidates.Add(new ProjectionTargetSpan(group, start, end));
        }
        if (candidates.Count == 0)
        {
            failure = "template-migration-semantic-value-baseline-empty";
            return false;
        }
        if (candidates.Count != 1)
        {
            failure = "template-migration-semantic-value-baseline-value-ambiguous";
            return false;
        }
        target = candidates[0];
        return true;
    }

    private static IReadOnlyList<IReadOnlyList<TemplateMigrationObject>> ProjectionRunGroups(IReadOnlyList<TemplateMigrationObject> runs)
        => runs.GroupBy(run => Regex.Replace(run.Id, ":run:\\d+$", string.Empty, RegexOptions.CultureInvariant), StringComparer.Ordinal)
            .Select(group => (IReadOnlyList<TemplateMigrationObject>)group.ToList())
            .ToList();

    private static bool ProjectionValueKindMatches(string value, string valueKind)
        => valueKind switch
        {
            "text" => value.Length != 0,
            "token" => Regex.IsMatch(value, "^\\S+$", RegexOptions.CultureInvariant),
            "date" => IsValidProjectionDate(value),
            "identifier" => value.Length <= 128 && value.Any(char.IsLetter) && value.All(character => char.IsLetterOrDigit(character) || character is '.' or '_' or '/' or '-' or '－' or '—'),
            "version" => Regex.IsMatch(value, "^(?:[0-9]{2}|[0-9]+\\.[0-9]+)$", RegexOptions.CultureInvariant),
            _ => false
        };

    private static bool IsValidProjectionDate(string value)
    {
        var formats = new[] { "yyyy-M-d", "yyyy/M/d", "yyyy.M.d", "yyyy年M月d日" };
        return DateOnly.TryParseExact(value, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out _);
    }

    private static bool IsProjectionPlaceholder(string value)
        => Regex.IsMatch(value, "^(?:\\{\\{[^{}]+\\}\\}|\\[[^\\[\\]]+\\])$", RegexOptions.CultureInvariant);

    private sealed record DelimitedValueSpan(int Start, int End, string Value);

    private static IReadOnlyList<DelimitedValueSpan> DelimitedValueSpans(string text, string valueKind, bool allowPlaceholder)
    {
        var spans = new List<DelimitedValueSpan>();
        for (var index = 0; index < text.Length; index += 1)
        {
            if (text[index] is not (':' or '：')) continue;
            var start = index + 1;
            while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
            var end = start;
            if (allowPlaceholder && start + 1 < text.Length && text[start] == '{' && text[start + 1] == '{')
            {
                var close = text.IndexOf("}}", start + 2, StringComparison.Ordinal);
                end = close < 0 ? start : close + 2;
            }
            else if (allowPlaceholder && start < text.Length && text[start] == '[')
            {
                var close = text.IndexOf(']', start + 1);
                end = close < 0 ? start : close + 1;
            }
            else
            {
                while (end < text.Length && ProjectionValueCharacterAllowed(text[end], valueKind)) end += 1;
            }
            if (end > start) spans.Add(new DelimitedValueSpan(start, end, text[start..end]));
        }
        return spans;
    }

    private static bool ProjectionValueCharacterAllowed(char character, string valueKind)
        => valueKind switch
        {
            "version" => char.IsAsciiDigit(character) || character == '.',
            "identifier" => char.IsLetterOrDigit(character) || character is '.' or '_' or '/' or '-' or '－' or '—',
            "date" => char.IsAsciiDigit(character) || character is '-' or '/' or '.' or '年' or '月' or '日',
            "token" => !char.IsWhiteSpace(character) && character is not (':' or '：'),
            _ => false
        };

    private static int FirstDelimiter(string text)
    {
        var ascii = text.IndexOf(':', StringComparison.Ordinal);
        var fullWidth = text.IndexOf('：');
        if (ascii < 0) return fullWidth;
        if (fullWidth < 0) return ascii;
        return Math.Min(ascii, fullWidth);
    }

    private static IReadOnlyList<ProjectionRunReplacement> BuildProjectionRunReplacements(ProjectionTargetSpan target, string value)
    {
        var replacements = new List<ProjectionRunReplacement>();
        var offset = 0;
        var affected = new List<(TemplateMigrationObject Run, int Start, int End)>();
        foreach (var run in target.Runs)
        {
            var text = run.Text ?? string.Empty;
            var start = offset;
            var end = start + text.Length;
            if (end > target.ValueStart && start < target.ValueEnd) affected.Add((run, start, end));
            offset = end;
        }
        for (var index = 0; index < affected.Count; index += 1)
        {
            var (run, runStart, runEnd) = affected[index];
            var original = run.Text ?? string.Empty;
            var prefixLength = index == 0 ? Math.Max(0, target.ValueStart - runStart) : 0;
            var suffixStart = index == affected.Count - 1 ? Math.Min(original.Length, target.ValueEnd - runStart) : original.Length;
            var prefix = prefixLength == 0 ? string.Empty : original[..prefixLength];
            var suffix = suffixStart >= original.Length ? string.Empty : original[suffixStart..];
            replacements.Add(new ProjectionRunReplacement(run, index == 0 ? prefix + value + suffix : suffix));
        }
        return replacements;
    }

    private static bool TryIndependentlyDeriveProjectionValue(
        IReadOnlyList<TemplateMigrationObject> objects,
        string parentId,
        string valueKind,
        string extraction,
        out string value)
    {
        value = string.Empty;
        if (extraction is not ("after-first-delimiter" or "unique-delimited-run-group" or "unique-delimited-value" or "whole-parent")) return false;
        var runs = objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, parentId, StringComparison.Ordinal)).ToList();
        if (string.Equals(extraction, "whole-parent", StringComparison.Ordinal))
        {
            var observed = string.Concat(runs.Select(run => run.Text ?? string.Empty)).Trim();
            if (!IndependentlyMatchesValueKind(observed, valueKind)) return false;
            value = observed;
            return true;
        }
        if (string.Equals(extraction, "unique-delimited-value", StringComparison.Ordinal))
        {
            var delimitedCandidates = new List<string>();
            foreach (var group in runs.GroupBy(run => Regex.Replace(run.Id, ":run:[0-9]+$", string.Empty, RegexOptions.CultureInvariant), StringComparer.Ordinal))
            {
                var text = string.Concat(group.Select(run => run.Text ?? string.Empty));
                for (var delimiter = 0; delimiter < text.Length; delimiter += 1)
                {
                    if (text[delimiter] is not (':' or '：')) continue;
                    var start = delimiter + 1;
                    while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
                    var end = start;
                    while (end < text.Length && IndependentlyAllowedValueCharacter(text[end], valueKind)) end += 1;
                    if (end <= start) continue;
                    var candidate = text[start..end];
                    if (IndependentlyMatchesValueKind(candidate, valueKind)) delimitedCandidates.Add(candidate);
                }
            }
            if (delimitedCandidates.Count != 1) return false;
            value = delimitedCandidates[0];
            return true;
        }
        var groups = string.Equals(extraction, "unique-delimited-run-group", StringComparison.Ordinal)
            ? runs.GroupBy(run => Regex.Replace(run.Id, ":run:[0-9]+$", string.Empty, RegexOptions.CultureInvariant), StringComparer.Ordinal).Select(group => group.ToList()).ToList()
            : [runs];
        var candidates = new List<string>();
        foreach (var group in groups)
        {
            var observed = new StringBuilder();
            foreach (var run in group) observed.Append(run.Text ?? string.Empty);
            var text = observed.ToString();
            var delimiter = -1;
            for (var index = 0; index < text.Length; index += 1) if (text[index] is ':' or '：') { delimiter = index; break; }
            if (delimiter < 0) continue;
            var candidate = text[(delimiter + 1)..].Trim();
            var match = IndependentlyMatchesValueKind(candidate, valueKind);
            if (match) candidates.Add(candidate);
        }
        if (candidates.Count != 1) return false;
        value = candidates[0];
        return true;
    }

    private static bool TryIndependentlyBuildProjectionReplacements(
        IReadOnlyList<TemplateMigrationObject> objects,
        string parentId,
        string valueKind,
        string extraction,
        string value,
        out IReadOnlyDictionary<string, string> replacements)
    {
        replacements = new Dictionary<string, string>(StringComparer.Ordinal);
        if (extraction is not ("after-first-delimiter" or "unique-delimited-run-group" or "unique-delimited-value" or "whole-parent")) return false;
        var runs = objects.Where(item => item.Kind == "run" && string.Equals(item.ParentId, parentId, StringComparison.Ordinal)).ToList();
        if (string.Equals(extraction, "whole-parent", StringComparison.Ordinal))
        {
            var text = string.Concat(runs.Select(run => run.Text ?? string.Empty));
            var start = 0; while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
            var end = text.Length; while (end > start && char.IsWhiteSpace(text[end - 1])) end -= 1;
            var observed = start < end ? text[start..end] : string.Empty;
            if (!IndependentlyMatchesValueKind(observed, valueKind) && !Regex.IsMatch(observed, "^(?:\\{\\{[^{}]+\\}\\}|\\[[^\\[\\]]+\\])$", RegexOptions.CultureInvariant)) return false;
            var target = new ProjectionTargetSpan(runs, start, end);
            replacements = BuildProjectionRunReplacements(target, value).ToDictionary(item => item.Run.Id, item => item.Text, StringComparer.Ordinal);
            return true;
        }
        if (string.Equals(extraction, "unique-delimited-value", StringComparison.Ordinal))
        {
            var delimitedCandidates = new List<(IReadOnlyList<TemplateMigrationObject> Runs, int Start, int End)>();
            foreach (var group in runs.GroupBy(run => Regex.Replace(run.Id, ":run:[0-9]+$", string.Empty, RegexOptions.CultureInvariant), StringComparer.Ordinal).Select(group => group.ToList()))
            {
                var text = string.Concat(group.Select(run => run.Text ?? string.Empty));
                for (var delimiter = 0; delimiter < text.Length; delimiter += 1)
                {
                    if (text[delimiter] is not (':' or '：')) continue;
                    var start = delimiter + 1;
                    while (start < text.Length && char.IsWhiteSpace(text[start])) start += 1;
                    var end = start;
                    if (start + 1 < text.Length && text[start] == '{' && text[start + 1] == '{')
                    {
                        var close = text.IndexOf("}}", start + 2, StringComparison.Ordinal);
                        end = close < 0 ? start : close + 2;
                    }
                    else if (start < text.Length && text[start] == '[')
                    {
                        var close = text.IndexOf(']', start + 1);
                        end = close < 0 ? start : close + 1;
                    }
                    else
                    {
                        while (end < text.Length && IndependentlyAllowedValueCharacter(text[end], valueKind)) end += 1;
                    }
                    if (end <= start) continue;
                    var candidate = text[start..end];
                    if (IndependentlyMatchesValueKind(candidate, valueKind) || Regex.IsMatch(candidate, "^(?:\\{\\{[^{}]+\\}\\}|\\[[^\\[\\]]+\\])$", RegexOptions.CultureInvariant)) delimitedCandidates.Add((group, start, end));
                }
            }
            if (delimitedCandidates.Count != 1) return false;
            var delimitedSelected = delimitedCandidates[0];
            var delimitedChanges = new Dictionary<string, string>(StringComparer.Ordinal);
            var delimitedOffset = 0;
            var delimitedAffected = new List<(TemplateMigrationObject Run, int Start, int End)>();
            foreach (var run in delimitedSelected.Runs)
            {
                var runText = run.Text ?? string.Empty;
                var start = delimitedOffset;
                var end = start + runText.Length;
                if (end > delimitedSelected.Start && start < delimitedSelected.End) delimitedAffected.Add((run, start, end));
                delimitedOffset = end;
            }
            for (var index = 0; index < delimitedAffected.Count; index += 1)
            {
                var (run, runStart, _) = delimitedAffected[index];
                var original = run.Text ?? string.Empty;
                var prefixLength = index == 0 ? Math.Max(0, delimitedSelected.Start - runStart) : 0;
                var suffixStart = index == delimitedAffected.Count - 1 ? Math.Min(original.Length, delimitedSelected.End - runStart) : original.Length;
                delimitedChanges[run.Id] = (index == 0 ? original[..prefixLength] + value : string.Empty) + (suffixStart < original.Length ? original[suffixStart..] : string.Empty);
            }
            replacements = delimitedChanges;
            return true;
        }
        var groups = string.Equals(extraction, "unique-delimited-run-group", StringComparison.Ordinal)
            ? runs.GroupBy(run => Regex.Replace(run.Id, ":run:[0-9]+$", string.Empty, RegexOptions.CultureInvariant), StringComparer.Ordinal).Select(group => group.ToList()).ToList()
            : [runs];
        var candidates = new List<(IReadOnlyList<TemplateMigrationObject> Runs, int Start, int End)>();
        foreach (var group in groups)
        {
            var groupText = string.Concat(group.Select(run => run.Text ?? string.Empty));
            var delimiter = -1;
            for (var index = 0; index < groupText.Length; index += 1) if (groupText[index] is ':' or '：') { delimiter = index; break; }
            if (delimiter < 0) continue;
            var start = delimiter + 1;
            while (start < groupText.Length && char.IsWhiteSpace(groupText[start])) start += 1;
            var end = groupText.Length;
            while (end > start && char.IsWhiteSpace(groupText[end - 1])) end -= 1;
            if (start >= end) continue;
            var candidate = groupText[start..end];
            var match = IndependentlyMatchesValueKind(candidate, valueKind);
            var placeholder = Regex.IsMatch(candidate, "^(?:\\{\\{[^{}]+\\}\\}|\\[[^\\[\\]]+\\])$", RegexOptions.CultureInvariant);
            if (match || placeholder) candidates.Add((group, start, end));
        }
        if (candidates.Count != 1) return false;
        var selected = candidates[0];
        var changed = new Dictionary<string, string>(StringComparer.Ordinal);
        var offset = 0;
        var affected = new List<(TemplateMigrationObject Run, int Start, int End)>();
        foreach (var run in selected.Runs)
        {
            var runText = run.Text ?? string.Empty;
            var start = offset;
            var end = start + runText.Length;
            if (end > selected.Start && start < selected.End) affected.Add((run, start, end));
            offset = end;
        }
        for (var index = 0; index < affected.Count; index += 1)
        {
            var (run, runStart, _) = affected[index];
            var original = run.Text ?? string.Empty;
            var prefixLength = index == 0 ? Math.Max(0, selected.Start - runStart) : 0;
            var suffixStart = index == affected.Count - 1 ? Math.Min(original.Length, selected.End - runStart) : original.Length;
            var prefix = prefixLength == 0 ? string.Empty : original[..prefixLength];
            var suffix = suffixStart >= original.Length ? string.Empty : original[suffixStart..];
            changed[run.Id] = index == 0 ? prefix + value + suffix : suffix;
        }
        replacements = changed;
        return true;
    }

    private static bool IndependentlyMatchesValueKind(string value, string valueKind)
        => valueKind switch
        {
            "text" => value.Length != 0,
            "token" => value.Length != 0 && value.All(character => !char.IsWhiteSpace(character)),
            "date" => DateOnly.TryParseExact(value, new[] { "yyyy-M-d", "yyyy/M/d", "yyyy.M.d", "yyyy年M月d日" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out _),
            "identifier" => value.Length is > 0 and <= 128 && value.Any(char.IsLetter) && value.All(character => char.IsLetterOrDigit(character) || character is '.' or '_' or '/' or '-' or '－' or '—'),
            "version" => Regex.IsMatch(value, "^(?:[0-9]{2}|[0-9]+\\.[0-9]+)$", RegexOptions.CultureInvariant),
            _ => false
        };

    private static bool IndependentlyAllowedValueCharacter(char character, string valueKind)
        => valueKind switch
        {
            "version" => character is >= '0' and <= '9' || character == '.',
            "identifier" => char.IsLetterOrDigit(character) || character is '.' or '_' or '/' or '-' or '－' or '—',
            "date" => character is >= '0' and <= '9' || character is '-' or '/' or '.' or '年' or '月' or '日',
            "token" => !char.IsWhiteSpace(character) && character is not (':' or '：'),
            _ => false
        };

    private static string HashCanonical<T>(T value)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value, Json.Options))));
    }

    private static string StructureFingerprint(TemplateMigrationObject item)
        => string.Join("|",
            item.Id,
            item.Kind,
            item.Scope,
            item.ParentId ?? string.Empty,
            item.Style ?? string.Empty,
            item.Kind == "section" ? item.Provenance.GetValueOrDefault("sectionPropertiesSha256") : string.Empty,
            item.Kind == "section" ? item.Provenance.GetValueOrDefault("headerFooterReferencesSha256") : string.Empty);

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }
}
