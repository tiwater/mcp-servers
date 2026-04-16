using System.Text.Json;

namespace Dockit.Docx;

public static class Resolver
{
    private const string SupportedScenario = "stability-report";

    public static int RunResolve(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("resolve requires <plan.json> <resolve-data.json>");
        }

        var planPath = Path.GetFullPath(args[0]);
        var dataPath = Path.GetFullPath(args[1]);
        if (!File.Exists(planPath))
        {
            throw new InvalidOperationException($"Plan file not found: {planPath}");
        }

        if (!File.Exists(dataPath))
        {
            throw new InvalidOperationException($"Resolve data file not found: {dataPath}");
        }

        var plan = JsonSerializer.Deserialize<DocxPlanResult>(File.ReadAllText(planPath), Json.Options)
            ?? throw new InvalidOperationException("Could not parse plan JSON.");
        var request = JsonSerializer.Deserialize<DocxResolveRequest>(File.ReadAllText(dataPath), Json.Options)
            ?? new DocxResolveRequest(null, null, null, null, null, null, null);
        var result = Resolve(plan, request);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static DocxResolveResult Resolve(DocxPlanResult plan, DocxResolveRequest request)
    {
        var scenario = string.IsNullOrWhiteSpace(request.Scenario) ? plan.Scenario : request.Scenario.Trim();
        if (!scenario.StartsWith(SupportedScenario, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Unsupported resolve scenario: {scenario}");
        }

        var stabilityRows = LoadStabilityRows(request.StabilityDataPath);
        var protocolMap = LoadProtocolValueMap(request.ProtocolPath);
        var qualityMap = LoadQualityStandardMap(request.QualityStandardCnPath);
        var reportTables = LoadReportTables(request.ReportPath);
        var inspectionMap = LoadInspectionReportValueMap(request.InspectionReportPath);
        var samplingPlanRows = LoadSamplingPlanRows(request.SamplingPlanPath);

        var operations = new List<DocxEditOperation>();
        var resolved = new HashSet<string>(StringComparer.Ordinal);
        var unresolved = new List<DocxResolveUnresolvedItem>();
        var warnings = new List<string>();

        foreach (var item in plan.Items)
        {
            if (item.ProposedEdits.Count > 0)
            {
                operations.AddRange(item.ProposedEdits);
                resolved.Add(item.CommentId);
                continue;
            }

            switch (item.CommentId)
            {
                case "0":
                    ResolveComment0(item, stabilityRows, operations, resolved, unresolved);
                    break;
                case "18":
                    ResolveComment18(item, stabilityRows, operations, resolved, unresolved);
                    break;
                case "20":
                    ResolveComment20(item, stabilityRows, operations, resolved, unresolved);
                    break;
                case "22":
                    ResolveComment22(item, stabilityRows, operations, resolved, unresolved);
                    break;
                case "14":
                    ResolveResultsTableLabels(item, reportTables, operations, resolved, unresolved);
                    break;
                case "15":
                    ResolveResultsTableCriteria(item, reportTables, qualityMap, operations, resolved, unresolved);
                    break;
                case "16":
                    ResolveResultsTableData(item, reportTables, stabilityRows, operations, resolved, unresolved);
                    break;
                case "8":
                    ResolveSampleInfoTable(
                        item,
                        reportTables,
                        stabilityRows,
                        inspectionMap,
                        operations,
                        resolved,
                        unresolved);
                    break;
                case "10":
                    ResolveSampleInfoStrength(
                        item,
                        reportTables,
                        protocolMap,
                        operations,
                        resolved,
                        unresolved);
                    break;
                case "1":
                    ResolveSamplingScheduleColumn(item, reportTables, samplingPlanRows, 2, 0, operations, resolved, unresolved);
                    break;
                case "2":
                    ResolveSamplingScheduleColumn(item, reportTables, samplingPlanRows, 3, 6, operations, resolved, unresolved);
                    break;
                case "4":
                    ResolveSamplingScheduleColumn(item, reportTables, samplingPlanRows, 4, 7, operations, resolved, unresolved);
                    break;
                default:
                    unresolved.Add(new DocxResolveUnresolvedItem(
                        item.CommentId,
                        item.InstructionType,
                        "The first ANA03 resolve slice does not handle this plan item yet."));
                    break;
            }
        }

        if (operations.Count == 0)
        {
            warnings.Add("No explicit operations were resolved from the provided plan.");
        }

        return new DocxResolveResult(
            Path.GetFullPath(plan.Input),
            scenario,
            operations,
            resolved.OrderBy(id => id, StringComparer.Ordinal).ToList(),
            unresolved,
            warnings);
    }

    private static void ResolveComment0(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var name = GetPrimaryValue(stabilityRows, "名称");
        var batch = GetPrimaryValue(stabilityRows, "批号");
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(batch))
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not resolve project name and batch from the stability workbook export."));
            return;
        }

        operations.Add(new DocxEditOperation(
            "replaceAnchoredText",
            CommentId: item.CommentId,
            Text: $"{name}工程批原液（{batch}）"));
        resolved.Add(item.CommentId);
    }

    private static void ResolveSampleInfoTable(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        IReadOnlyDictionary<string, string> inspectionMap,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var tableIndex = item.Anchor.TableIndex
            ?? FindTableIndex(reportTables, row => row.Count > 0 && row[0].Contains("产品名称", StringComparison.Ordinal));
        if (tableIndex is null || tableIndex < 0 || tableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the sample-info table in the target report."));
            return;
        }

        var rows = reportTables[tableIndex.Value];
        var name = GetPrimaryValue(stabilityRows, "名称");
        if (string.IsNullOrWhiteSpace(name))
        {
            name = GetMappedValue(inspectionMap, "产品名称");
        }

        var batch = GetPrimaryValue(stabilityRows, "批号");
        if (string.IsNullOrWhiteSpace(batch))
        {
            batch = GetMappedValue(inspectionMap, "批号");
        }

        var rowValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["产品名称"] = name,
            ["生产日期"] = GetMappedValue(inspectionMap, "生产日期"),
            ["生产厂家"] = GetMappedValue(inspectionMap, "生产厂家"),
            ["总数量"] = GetMappedValue(inspectionMap, "总数量"),
            ["批号"] = batch,
        };

        var count = 0;
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var firstCell = rows[rowIndex].Count > 0 ? rows[rowIndex][0] : string.Empty;
            var key = NormalizeSampleInfoKey(firstCell);
            if (string.IsNullOrWhiteSpace(key) || !rowValues.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            operations.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: tableIndex,
                RowIndex: rowIndex,
                CellIndex: 1,
                Text: value));
            count++;
        }

        if (count == 0)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not map any sample-info fields from the inspection report and workbook exports."));
            return;
        }

        resolved.Add(item.CommentId);
    }

    private static void ResolveSampleInfoStrength(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        IReadOnlyDictionary<string, string> protocolMap,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var tableIndex = item.Anchor.TableIndex
            ?? FindTableIndex(reportTables, row => row.Count > 0 && row[0].Contains("规格", StringComparison.Ordinal));
        if (tableIndex is null || tableIndex < 0 || tableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the sample-info table for the strength field."));
            return;
        }

        var strength = GetMappedValue(protocolMap, "分装规格");
        if (string.IsNullOrWhiteSpace(strength))
        {
            strength = GetMappedValue(protocolMap, "规格");
        }

        if (string.IsNullOrWhiteSpace(strength))
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not resolve the strength field from the stability protocol export."));
            return;
        }

        var rows = reportTables[tableIndex.Value];
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var firstCell = rows[rowIndex].Count > 0 ? rows[rowIndex][0] : string.Empty;
            if (!NormalizeSampleInfoKey(firstCell).Equals("规格", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            operations.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: tableIndex,
                RowIndex: rowIndex,
                CellIndex: 1,
                Text: strength));
            resolved.Add(item.CommentId);
            return;
        }

        unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not find the strength row inside the sample-info table."));
    }

    private static void ResolveSamplingScheduleColumn(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        IReadOnlyList<IReadOnlyList<string>> samplingPlanRows,
        int targetCellIndex,
        int sourceColumnIndex,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        if (item.Anchor.TableIndex is null || item.Anchor.TableIndex < 0 || item.Anchor.TableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the target sampling-schedule table in the report."));
            return;
        }

        var tableRows = reportTables[item.Anchor.TableIndex.Value];
        var targetRows = GetSamplingScheduleTargetRows(tableRows);
        if (targetRows.Count == 0)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the target schedule rows in the report table."));
            return;
        }

        var sourceRows = GetSamplingScheduleSourceRows(samplingPlanRows);
        if (sourceRows.Count < targetRows.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "The exported sampling-plan source does not contain enough influence-factor rows to fill the report schedule table."));
            return;
        }

        for (var i = 0; i < targetRows.Count; i++)
        {
            var rawValue = GetCellValue(sourceRows[i], sourceColumnIndex);
            operations.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: item.Anchor.TableIndex,
                RowIndex: targetRows[i],
                CellIndex: targetCellIndex,
                Text: NormalizeSamplingScheduleValue(rawValue)));
        }

        resolved.Add(item.CommentId);
    }

    private static void ResolveComment18(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var name = GetPrimaryValue(stabilityRows, "名称");
        var batch = GetPrimaryValue(stabilityRows, "批号");
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(batch) || item.Anchor.ParagraphIndex is null || string.IsNullOrWhiteSpace(item.Anchor.CurrentParagraphText))
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not resolve paragraph replacement inputs for the condition-summary paragraph."));
            return;
        }

        var text = item.Anchor.CurrentParagraphText
            .Replace("HSPXXXXDS", name, StringComparison.Ordinal)
            .Replace("YYYY", batch, StringComparison.Ordinal);

        operations.Add(new DocxEditOperation(
            "replaceParagraphText",
            ParagraphIndex: item.Anchor.ParagraphIndex,
            Text: text));
        resolved.Add(item.CommentId);
    }

    private static void ResolveComment20(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var text = ReplaceScenarioPlaceholders(item.Anchor.FollowingParagraphText, stabilityRows);
        if (string.IsNullOrWhiteSpace(text) || item.Anchor.ParagraphIndex is null)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not resolve the photostability introduction paragraph from the existing annotated draft."));
            return;
        }

        operations.Add(new DocxEditOperation(
            "replaceParagraphText",
            ParagraphIndex: item.Anchor.ParagraphIndex + 1,
            Text: text));
        resolved.Add(item.CommentId);
    }

    private static void ResolveComment22(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        var text = ReplaceScenarioPlaceholders(item.Anchor.CurrentParagraphText, stabilityRows);
        if (string.IsNullOrWhiteSpace(text) || item.Anchor.ParagraphIndex is null)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not resolve the existing summary paragraph draft."));
            return;
        }

        operations.Add(new DocxEditOperation(
            "replaceParagraphText",
            ParagraphIndex: item.Anchor.ParagraphIndex,
            Text: text));
        resolved.Add(item.CommentId);
    }

    private static void ResolveResultsTableLabels(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        if (item.Anchor.TableIndex is null || item.Anchor.TableIndex < 0 || item.Anchor.TableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the target results table in the report."));
            return;
        }

        var tableRows = reportTables[item.Anchor.TableIndex.Value];
        var mappings = GetResolvableResultRows(tableRows);
        if (mappings.Count == 0)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "No resolvable result rows were found in the target table."));
            return;
        }

        foreach (var mapping in mappings)
        {
            operations.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: item.Anchor.TableIndex,
                RowIndex: mapping.RowIndex,
                CellIndex: 0,
                Text: mapping.DisplayLabel));
        }

        resolved.Add(item.CommentId);
    }

    private static void ResolveResultsTableCriteria(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        IReadOnlyDictionary<string, string> qualityMap,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        if (item.Anchor.TableIndex is null || item.Anchor.TableIndex < 0 || item.Anchor.TableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the target results table in the report."));
            return;
        }

        var tableRows = reportTables[item.Anchor.TableIndex.Value];
        var mappings = GetResolvableResultRows(tableRows);
        var count = 0;
        foreach (var mapping in mappings)
        {
            if (!qualityMap.TryGetValue(mapping.Key, out var criteria))
            {
                continue;
            }

            operations.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: item.Anchor.TableIndex,
                RowIndex: mapping.RowIndex,
                CellIndex: 1,
                Text: criteria));
            count++;
        }

        if (count == 0)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not map any acceptance criteria from the Chinese quality standard export."));
            return;
        }

        resolved.Add(item.CommentId);
    }

    private static void ResolveResultsTableData(
        DocxPlanItem item,
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> reportTables,
        IReadOnlyList<IReadOnlyList<string>> stabilityRows,
        List<DocxEditOperation> operations,
        HashSet<string> resolved,
        List<DocxResolveUnresolvedItem> unresolved)
    {
        if (item.Anchor.TableIndex is null || item.Anchor.TableIndex < 0 || item.Anchor.TableIndex >= reportTables.Count)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not locate the target results table in the report."));
            return;
        }

        var tableRows = reportTables[item.Anchor.TableIndex.Value];
        var mappings = GetResolvableResultRows(tableRows);
        var count = 0;
        foreach (var mapping in mappings)
        {
            if (!TryGetStabilityRow(stabilityRows, mapping.WorkbookRowLabel, out var workbookRow))
            {
                continue;
            }

            var values = new[]
            {
                GetCellValue(workbookRow, 4),
                GetCellValue(workbookRow, 12),
                GetCellValue(workbookRow, 13),
            };

            for (var i = 0; i < values.Length; i++)
            {
                operations.Add(new DocxEditOperation(
                    "replaceTableCellText",
                    TableIndex: item.Anchor.TableIndex,
                    RowIndex: mapping.RowIndex,
                    CellIndex: 2 + i,
                    Text: values[i]));
            }

            count++;
        }

        if (count == 0)
        {
            unresolved.Add(new DocxResolveUnresolvedItem(item.CommentId, item.InstructionType, "Could not map any freeze-thaw result rows from the stability workbook export."));
            return;
        }

        resolved.Add(item.CommentId);
    }

    private static IReadOnlyList<IReadOnlyList<string>> LoadStabilityRows(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException("stabilityDataPath is required and must point to an exported ANA03 workbook JSON file.");
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        if (root.ValueKind != JsonValueKind.Array || root.GetArrayLength() == 0)
        {
            throw new InvalidOperationException("The ANA03 workbook export JSON must be a non-empty array.");
        }

        var firstSheet = root[0];
        if (!firstSheet.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("The ANA03 workbook export JSON must contain a rows array.");
        }

        var rows = new List<IReadOnlyList<string>>();
        foreach (var row in rowsElement.EnumerateArray())
        {
            var cells = new List<string>();
            foreach (var cell in row.EnumerateArray())
            {
                cells.Add(cell.GetString() ?? string.Empty);
            }
            rows.Add(cells);
        }

        return rows;
    }

    private static IReadOnlyList<IReadOnlyList<string>> LoadSamplingPlanRows(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return [];
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        if (root.ValueKind != JsonValueKind.Array || root.GetArrayLength() == 0)
        {
            return [];
        }

        var firstSheet = root[0];
        if (!firstSheet.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var rows = new List<IReadOnlyList<string>>();
        foreach (var row in rowsElement.EnumerateArray())
        {
            var cells = new List<string>();
            foreach (var cell in row.EnumerateArray())
            {
                cells.Add(cell.GetString() ?? string.Empty);
            }
            rows.Add(cells);
        }

        return rows;
    }

    private static IReadOnlyDictionary<string, string> LoadProtocolValueMap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.ValueKind != JsonValueKind.Array)
        {
            return map;
        }

        foreach (var node in root.EnumerateArray())
        {
            if (!node.TryGetProperty("type", out var typeElement) || !string.Equals(typeElement.GetString(), "table", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!node.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var row in rowsElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() < 2)
                {
                    continue;
                }

                for (var i = 0; i + 1 < row.GetArrayLength(); i += 2)
                {
                    var key = row[i].GetString() ?? string.Empty;
                    var value = row[i + 1].GetString() ?? string.Empty;
                    var normalizedKey = NormalizeSampleInfoKey(key);
                    if (!string.IsNullOrWhiteSpace(normalizedKey) && !string.IsNullOrWhiteSpace(value) && !map.ContainsKey(normalizedKey))
                    {
                        map[normalizedKey] = value.Trim();
                    }
                }
            }
        }

        return map;
    }

    private static IReadOnlyDictionary<string, string> LoadInspectionReportValueMap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.ValueKind != JsonValueKind.Array)
        {
            return map;
        }

        foreach (var node in root.EnumerateArray())
        {
            if (!node.TryGetProperty("type", out var typeElement) || !string.Equals(typeElement.GetString(), "table", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!node.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var row in rowsElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() < 2)
                {
                    continue;
                }

                for (var i = 0; i + 1 < row.GetArrayLength(); i += 2)
                {
                    var key = row[i].GetString() ?? string.Empty;
                    var value = row[i + 1].GetString() ?? string.Empty;
                    var normalizedKey = NormalizeSampleInfoKey(key);
                    if (!string.IsNullOrWhiteSpace(normalizedKey) && !string.IsNullOrWhiteSpace(value) && !map.ContainsKey(normalizedKey))
                    {
                        map[normalizedKey] = value.Trim();
                    }
                }
            }
        }

        return map;
    }

    private static IReadOnlyDictionary<string, string> LoadQualityStandardMap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException("qualityStandardCnPath is required and must point to an exported Chinese quality-standard JSON file.");
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.ValueKind != JsonValueKind.Array)
        {
            return map;
        }

        foreach (var node in root.EnumerateArray())
        {
            if (!node.TryGetProperty("type", out var typeElement) || !string.Equals(typeElement.GetString(), "table", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!node.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var row in rowsElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() < 3)
                {
                    continue;
                }

                var second = row[1].GetString() ?? string.Empty;
                var criteria = row[2].GetString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(second) || string.IsNullOrWhiteSpace(criteria))
                {
                    continue;
                }

                var key = NormalizeAnalyteKey(second);
                if (!string.IsNullOrWhiteSpace(key) && !map.ContainsKey(key))
                {
                    map[key] = criteria;
                }
            }
        }

        return map;
    }

    private static IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> LoadReportTables(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException("reportPath is required and must point to an exported ANA03 report JSON file.");
        }

        var root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(Path.GetFullPath(path)), Json.Options);
        if (root.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("The ANA03 report export JSON must be an array of blocks.");
        }

        var tables = new List<IReadOnlyList<IReadOnlyList<string>>>();
        foreach (var node in root.EnumerateArray())
        {
            if (!node.TryGetProperty("type", out var typeElement) || !string.Equals(typeElement.GetString(), "table", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!node.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            var rows = new List<IReadOnlyList<string>>();
            foreach (var row in rowsElement.EnumerateArray())
            {
                var cells = new List<string>();
                foreach (var cell in row.EnumerateArray())
                {
                    cells.Add(cell.GetString() ?? string.Empty);
                }
                rows.Add(cells);
            }

            tables.Add(rows);
        }

        return tables;
    }

    private static List<int> GetSamplingScheduleTargetRows(IReadOnlyList<IReadOnlyList<string>> tableRows)
    {
        var result = new List<int>();
        for (var rowIndex = 2; rowIndex < tableRows.Count; rowIndex++)
        {
            if (tableRows[rowIndex].Any(cell => !string.IsNullOrWhiteSpace(cell)))
            {
                result.Add(rowIndex);
            }
        }

        return result;
    }

    private static List<IReadOnlyList<string>> GetSamplingScheduleSourceRows(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        return rows
            .Skip(1)
            .Where(row =>
                !string.IsNullOrWhiteSpace(GetCellValue(row, 3)) &&
                (GetCellValue(row, 3).Contains("高温", StringComparison.Ordinal) ||
                 GetCellValue(row, 3).Contains("光照", StringComparison.Ordinal) ||
                 GetCellValue(row, 3).Contains("冻融", StringComparison.Ordinal)))
            .Take(9)
            .ToList();
    }

    private static string GetPrimaryValue(IReadOnlyList<IReadOnlyList<string>> rows, string rowLabel)
    {
        if (!TryGetStabilityRow(rows, rowLabel, out var row))
        {
            return string.Empty;
        }

        for (var i = 0; i < row.Count; i++)
        {
            var value = row[i];
            if (!string.IsNullOrWhiteSpace(value) && i >= 5)
            {
                return value.Trim();
            }
        }

        return row.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;
    }

    private static bool TryGetStabilityRow(IReadOnlyList<IReadOnlyList<string>> rows, string rowLabel, out IReadOnlyList<string> row)
    {
        row = rows.FirstOrDefault(candidate => candidate.Count > 0 && string.Equals(candidate[0], rowLabel, StringComparison.OrdinalIgnoreCase))
            ?? [];
        return row.Count > 0;
    }

    private static string GetCellValue(IReadOnlyList<string> row, int index)
        => index >= 0 && index < row.Count ? row[index] : string.Empty;

    private static string NormalizeSamplingScheduleValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim().Replace('.', '-');
        var match = System.Text.RegularExpressions.Regex.Match(normalized, @"\d{4}-\d{2}-\d{2}");
        return match.Success ? match.Value : normalized;
    }

    private static string ReplaceScenarioPlaceholders(string? text, IReadOnlyList<IReadOnlyList<string>> stabilityRows)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var name = GetPrimaryValue(stabilityRows, "名称");
        var batch = GetPrimaryValue(stabilityRows, "批号");
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(batch))
        {
            return string.Empty;
        }

        var projectCode = name.EndsWith("DS", StringComparison.OrdinalIgnoreCase)
            ? name[..^2]
            : name;

        return text
            .Replace("HSPXXXXDS", name, StringComparison.Ordinal)
            .Replace("HSPXXXX", projectCode, StringComparison.Ordinal)
            .Replace("YYYY", batch, StringComparison.Ordinal);
    }

    private static string GetMappedValue(IReadOnlyDictionary<string, string> map, string key)
        => map.TryGetValue(key, out var value) ? value : string.Empty;

    private static int? FindTableIndex(
        IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> tables,
        Func<IReadOnlyList<string>, bool> rowPredicate)
    {
        for (var tableIndex = 0; tableIndex < tables.Count; tableIndex++)
        {
            if (tables[tableIndex].Any(rowPredicate))
            {
                return tableIndex;
            }
        }

        return null;
    }

    private static IReadOnlyList<ResultRowMapping> GetResolvableResultRows(IReadOnlyList<IReadOnlyList<string>> tableRows)
    {
        var mappings = new List<ResultRowMapping>();
        for (var rowIndex = 0; rowIndex < tableRows.Count; rowIndex++)
        {
            if (tableRows[rowIndex].Count == 0)
            {
                continue;
            }

            var firstCell = tableRows[rowIndex][0];
            var key = NormalizeAnalyteKey(firstCell);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            if (key.Equals("颜色", StringComparison.OrdinalIgnoreCase))
            {
                mappings.Add(new ResultRowMapping(rowIndex, key, "颜色", "颜色Color"));
            }
            else if (key.Equals("澄清度", StringComparison.OrdinalIgnoreCase))
            {
                mappings.Add(new ResultRowMapping(rowIndex, key, "澄清度", "澄清度Clarity"));
            }
            else if (key.Equals("ph", StringComparison.OrdinalIgnoreCase))
            {
                mappings.Add(new ResultRowMapping(rowIndex, key, "PH", "pH"));
            }
            else if (key.Equals("蛋白质含量", StringComparison.OrdinalIgnoreCase))
            {
                mappings.Add(new ResultRowMapping(rowIndex, key, "蛋白质含量", "蛋白质含量Protein concentration"));
            }
        }

        return mappings;
    }

    private static string NormalizeAnalyteKey(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var normalized = text
            .Replace("Protein concentration", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Clarity", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Color", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Test item", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Test Item", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Acceptance criteria", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Analytical Method", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" ", string.Empty)
            .Replace(" ", string.Empty)
            .Trim();

        if (normalized.Contains("颜色", StringComparison.Ordinal)) return "颜色";
        if (normalized.Contains("澄清度", StringComparison.Ordinal)) return "澄清度";
        if (normalized.Equals("pH", StringComparison.OrdinalIgnoreCase)) return "pH";
        if (normalized.Contains("蛋白质含量", StringComparison.Ordinal)) return "蛋白质含量";
        return normalized;
    }

    private static string NormalizeSampleInfoKey(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var normalized = text
            .Replace("Product Name", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Product name", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Production Date", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Date of Manufacture", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Manufacturer", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Batch Size", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Batch No.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Batch No", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Strength", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Specification No.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Target Concentration", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Storage Condition", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Total Quantity", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("贮存条件", string.Empty, StringComparison.Ordinal)
            .Replace("产品名称", string.Empty, StringComparison.Ordinal)
            .Replace("生产日期", string.Empty, StringComparison.Ordinal)
            .Replace("生产厂家", string.Empty, StringComparison.Ordinal)
            .Replace("批量", string.Empty, StringComparison.Ordinal)
            .Replace("批号", string.Empty, StringComparison.Ordinal)
            .Replace("规格", string.Empty, StringComparison.Ordinal)
            .Replace("分装规格", string.Empty, StringComparison.Ordinal)
            .Replace("总数量", string.Empty, StringComparison.Ordinal)
            .Replace("质量标准", string.Empty, StringComparison.Ordinal)
            .Replace("Specification", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("No.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" ", string.Empty)
            .Replace(" ", string.Empty)
            .Trim();

        if (text.Contains("产品名称", StringComparison.Ordinal) || text.Contains("Product name", StringComparison.OrdinalIgnoreCase) || text.Contains("Product Name", StringComparison.OrdinalIgnoreCase)) return "产品名称";
        if (text.Contains("生产日期", StringComparison.Ordinal) || text.Contains("Production Date", StringComparison.OrdinalIgnoreCase) || text.Contains("Date of Manufacture", StringComparison.OrdinalIgnoreCase)) return "生产日期";
        if (text.Contains("生产厂家", StringComparison.Ordinal) || text.Contains("Manufacturer", StringComparison.OrdinalIgnoreCase)) return "生产厂家";
        if (text.Contains("批量", StringComparison.Ordinal) || text.Contains("Batch Size", StringComparison.OrdinalIgnoreCase) || text.Contains("Total Quantity", StringComparison.OrdinalIgnoreCase) || text.Contains("总数量", StringComparison.Ordinal)) return "总数量";
        if (text.Contains("批号", StringComparison.Ordinal) || text.Contains("Batch No", StringComparison.OrdinalIgnoreCase)) return "批号";
        if (text.Contains("分装规格", StringComparison.Ordinal) || text.Contains("规格", StringComparison.Ordinal) || text.Contains("Strength", StringComparison.OrdinalIgnoreCase)) return text.Contains("分装规格", StringComparison.Ordinal) ? "分装规格" : "规格";
        return normalized;
    }

    private sealed record ResultRowMapping(int RowIndex, string Key, string WorkbookRowLabel, string DisplayLabel);
}
