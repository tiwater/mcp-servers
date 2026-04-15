using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Planner
{
    private const string Ana14Scenario = "experimental-record-attachment";

    public static int RunPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("plan requires <input.xlsx> <plan-data.json>");
        }

        var input = Path.GetFullPath(args[0]);
        var dataPath = Path.GetFullPath(args[1]);
        if (!File.Exists(dataPath))
        {
            throw new InvalidOperationException($"Plan data file not found: {dataPath}");
        }

        var json = File.ReadAllText(dataPath);
        var request = JsonSerializer.Deserialize<XlsxPlanRequest>(json, Json.Options)
            ?? throw new InvalidOperationException("Failed to parse planning data");
        var result = Plan(input, request);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static XlsxPlanResult Plan(string workbookPath, XlsxPlanRequest request)
    {
        var scenario = string.IsNullOrWhiteSpace(request.Scenario) ? Ana14Scenario : request.Scenario.Trim();
        if (!string.Equals(scenario, Ana14Scenario, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Unsupported planning scenario: {scenario}");
        }

        if (request.Sources is null || request.Sources.Count == 0)
        {
            throw new InvalidOperationException("At least one planning source is required.");
        }

        var workbook = LoadWorkbookModel(workbookPath, request.Sheet);
        var section280 = workbook.Sections.SingleOrDefault(section =>
            section.Kind == WorkbookSectionKind.Source280)
            ?? throw new InvalidOperationException("Could not find the 280nm section in the workbook.");
        var section360 = workbook.Sections.SingleOrDefault(section =>
            section.Kind == WorkbookSectionKind.Source360)
            ?? throw new InvalidOperationException("Could not find the 360nm section in the workbook.");
        var formulaSection = workbook.Sections.SingleOrDefault(section => section.Kind == WorkbookSectionKind.FormulaDerived);

        var warnings = new List<string>();
        var selectedSources = new List<XlsxPlanSelectedSource>();
        var plannedSections = new List<XlsxPlanSection>();
        var proposedEdits = new List<XlsxEditOperation>();

        BuildSectionPlan(section280, request.Sources, selectedSources, plannedSections, proposedEdits, warnings);
        BuildSectionPlan(section360, request.Sources, selectedSources, plannedSections, proposedEdits, warnings);

        if (formulaSection is not null)
        {
            plannedSections.Add(new XlsxPlanSection(
                "formula-derived",
                formulaSection.DataStartCell,
                formulaSection.DataEndCell,
                FormulaDriven: true,
                Samples: formulaSection.SampleRows
                    .Select(sample => new XlsxPlanSampleRow(sample.SampleId, sample.RowIndex, []))
                    .ToList()));
        }

        var confidence = warnings.Count == 0 ? "high" : "medium";
        return new XlsxPlanResult(
            Path.GetFullPath(workbookPath),
            scenario,
            workbook.SheetName,
            selectedSources,
            plannedSections,
            proposedEdits,
            warnings,
            confidence);
    }

    private static void BuildSectionPlan(
        WorkbookSection section,
        IReadOnlyList<XlsxPlanSourceDocument> sources,
        List<XlsxPlanSelectedSource> selectedSources,
        List<XlsxPlanSection> plannedSections,
        List<XlsxEditOperation> proposedEdits,
        List<string> warnings)
    {
        var source = SelectSourceTable(section, sources)
            ?? throw new InvalidOperationException($"Could not find a summarized source table for section '{section.Title}'.");

        selectedSources.Add(new XlsxPlanSelectedSource(source.Document.Name, source.Document.File, source.Table.Title, source.Table.Page));

        var samplePlans = new List<XlsxPlanSampleRow>();
        var values = new List<IReadOnlyList<string>>();
        foreach (var sample in section.SampleRows)
        {
            if (!TryFindSampleValues(source.Table, sample.SampleId, section.Headers, out var rowValues))
            {
                warnings.Add($"Sample '{sample.SampleId}' was not found in source '{source.Document.Name}' table '{source.Table.Title ?? "untitled"}'.");
                continue;
            }

            samplePlans.Add(new XlsxPlanSampleRow(sample.SampleId, sample.RowIndex, rowValues));
            values.Add(rowValues);
        }

        if (values.Count > 0)
        {
            proposedEdits.Add(new XlsxEditOperation(
                "setRangeValues",
                Sheet: section.SheetName,
                StartCell: section.DataStartCell,
                Values: values));
        }

        plannedSections.Add(new XlsxPlanSection(
            section.Title,
            section.DataStartCell,
            section.DataEndCell,
            FormulaDriven: false,
            Samples: samplePlans));
    }

    private static bool TryFindSampleValues(
        XlsxPlanSourceTable table,
        string sampleId,
        IReadOnlyList<string> expectedHeaders,
        out IReadOnlyList<string> rowValues)
    {
        rowValues = [];
        if (table.Header is null || table.Header.Count == 0)
        {
            return false;
        }

        var normalizedHeaderIndex = table.Header
            .Select((value, index) => (Key: NormalizeToken(value), Index: index))
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Key))
            .GroupBy(entry => entry.Key)
            .ToDictionary(group => group.Key, group => group.First().Index, StringComparer.Ordinal);

        var sampleColumnIndex = normalizedHeaderIndex.GetValueOrDefault("SAMPLENAME", -1);
        if (sampleColumnIndex < 0)
        {
            return false;
        }

        var matchingRow = table.Rows
            .Skip(1)
            .FirstOrDefault(row => sampleColumnIndex < row.Count &&
                string.Equals(row[sampleColumnIndex].Trim(), sampleId, StringComparison.OrdinalIgnoreCase));

        if (matchingRow is null)
        {
            return false;
        }

        var values = new List<string>();
        foreach (var expectedHeader in expectedHeaders)
        {
            var key = NormalizeToken(expectedHeader);
            if (!normalizedHeaderIndex.TryGetValue(key, out var columnIndex) || columnIndex >= matchingRow.Count)
            {
                return false;
            }

            values.Add(NormalizePlannedValue(matchingRow[columnIndex]));
        }

        rowValues = values;
        return true;
    }

    private static (XlsxPlanSourceDocument Document, XlsxPlanSourceTable Table)? SelectSourceTable(
        WorkbookSection section,
        IReadOnlyList<XlsxPlanSourceDocument> sources)
    {
        var sectionKey = section.Title.Contains("360", StringComparison.Ordinal) ? "360" : "280";

        foreach (var source in sources)
        {
            var sourceName = $"{source.Name} {source.File}".ToUpperInvariant();
            if (!sourceName.Contains(sectionKey, StringComparison.Ordinal))
            {
                continue;
            }

            var table = source.Tables.FirstOrDefault(candidate =>
                (candidate.Title?.Contains("Area Summarized by Name", StringComparison.OrdinalIgnoreCase) ?? false) &&
                candidate.Header is not null &&
                candidate.Header.Any(header => NormalizeToken(header) == "SAMPLENAME"));

            if (table is not null)
            {
                return (source, table);
            }
        }

        return null;
    }

    private static WorkbookModel LoadWorkbookModel(string workbookPath, string? requestedSheet)
    {
        using var spreadsheet = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part not found.");
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(candidate => string.IsNullOrWhiteSpace(requestedSheet)
                || string.Equals(candidate.Name?.Value, requestedSheet, StringComparison.Ordinal))
            ?? throw new InvalidOperationException("Could not find the requested sheet.");

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Sheet data not found.");

        var rowCells = sheetData.Elements<Row>()
            .ToDictionary(
                row => (int)(row.RowIndex?.Value ?? 0),
                row => row.Elements<Cell>().ToDictionary(
                    cell => cell.CellReference?.Value ?? string.Empty,
                    cell => GetCellValue(cell, sharedStrings),
                    StringComparer.OrdinalIgnoreCase));

        var sections = new List<WorkbookSection>();
        var orderedRows = rowCells.Keys.OrderBy(value => value).ToList();
        for (var index = 0; index < orderedRows.Count; index++)
        {
            var rowIndex = orderedRows[index];
            var title = GetValue(rowCells, rowIndex, "D");
            if (string.IsNullOrWhiteSpace(title))
            {
                continue;
            }

            if (!(title.Contains("280", StringComparison.Ordinal)
                || title.Contains("360", StringComparison.Ordinal)
                || title.Contains("*0.784", StringComparison.Ordinal)))
            {
                continue;
            }

            var headerRowIndex = rowIndex;
            var sampleRows = new List<WorkbookSampleRow>();
            var dataStartRow = rowIndex + 1;
            var cursor = dataStartRow;
            while (rowCells.TryGetValue(cursor, out _) && !string.IsNullOrWhiteSpace(GetValue(rowCells, cursor, "D")))
            {
                var sampleId = GetValue(rowCells, cursor, "D");
                if (GetSectionKind(sampleId) != WorkbookSectionKind.Unknown)
                {
                    break;
                }

                if (sampleId.Contains("注意", StringComparison.Ordinal))
                {
                    break;
                }

                sampleRows.Add(new WorkbookSampleRow(sampleId, cursor));
                cursor++;
            }

            var headers = ReadHeaders(rowCells, headerRowIndex);
            if (headers.Count == 0)
            {
                continue;
            }

            sections.Add(new WorkbookSection(
                sheet.Name?.Value ?? "Unknown",
                title,
                GetSectionKind(title),
                headerRowIndex,
                dataStartRow,
                DataEndRow: sampleRows.Count == 0 ? dataStartRow : sampleRows.Max(sample => sample.RowIndex),
                Headers: headers,
                SampleRows: sampleRows));
        }

        return new WorkbookModel(sheet.Name?.Value ?? "Unknown", sections);
    }

    private static List<string> ReadHeaders(
        IReadOnlyDictionary<int, Dictionary<string, string>> rowCells,
        int rowIndex)
    {
        var headers = new List<string>();
        for (var column = 5; column <= 11; column++)
        {
            var value = GetValue(rowCells, rowIndex, GetColumnReference(column));
            if (!string.IsNullOrWhiteSpace(value))
            {
                headers.Add(value);
            }
        }

        return headers;
    }

    private static string GetValue(IReadOnlyDictionary<int, Dictionary<string, string>> rowCells, int rowIndex, string column)
    {
        if (!rowCells.TryGetValue(rowIndex, out var cells))
        {
            return string.Empty;
        }

        var cellReference = $"{column}{rowIndex}";
        return cells.TryGetValue(cellReference, out var value) ? value : string.Empty;
    }

    private static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable is not null)
        {
            if (int.TryParse(cell.InnerText, out var sharedStringIndex))
            {
                return sharedStringTable.ElementAt(sharedStringIndex).InnerText;
            }
        }

        if (cell.InlineString?.Text?.Text is { } inlineText)
        {
            return inlineText;
        }

        return cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
    }

    private static string NormalizeToken(string value)
    {
        var cleaned = Regex.Replace(value.ToUpperInvariant(), "[^A-Z0-9]+", string.Empty);
        return cleaned;
    }

    private static string CleanCellText(string value)
        => value.Replace("\n", string.Empty, StringComparison.Ordinal).Trim();

    private static string NormalizePlannedValue(string value)
    {
        var cleaned = CleanCellText(value);
        return string.IsNullOrWhiteSpace(cleaned) ? "0" : cleaned;
    }

    private static string GetColumnReference(int columnIndex)
    {
        var letters = new Stack<char>();
        var index = columnIndex;
        while (index > 0)
        {
            index--;
            letters.Push((char)('A' + (index % 26)));
            index /= 26;
        }

        return new string(letters.ToArray());
    }

    private sealed record WorkbookModel(
        string SheetName,
        IReadOnlyList<WorkbookSection> Sections
    );

    private sealed record WorkbookSection(
        string SheetName,
        string Title,
        WorkbookSectionKind Kind,
        int HeaderRowIndex,
        int DataStartRow,
        int DataEndRow,
        IReadOnlyList<string> Headers,
        IReadOnlyList<WorkbookSampleRow> SampleRows)
    {
        public string DataStartCell => $"E{DataStartRow}";
        public string DataEndCell => $"K{DataEndRow}";
    }

    private sealed record WorkbookSampleRow(
        string SampleId,
        int RowIndex
    );

    private static WorkbookSectionKind GetSectionKind(string title)
    {
        if (title.Contains("*0.784", StringComparison.Ordinal))
        {
            return WorkbookSectionKind.FormulaDerived;
        }

        if (title.Contains("280", StringComparison.Ordinal))
        {
            return WorkbookSectionKind.Source280;
        }

        if (title.Contains("360", StringComparison.Ordinal))
        {
            return WorkbookSectionKind.Source360;
        }

        return WorkbookSectionKind.Unknown;
    }

    private enum WorkbookSectionKind
    {
        Unknown,
        Source280,
        Source360,
        FormulaDerived
    }
}
