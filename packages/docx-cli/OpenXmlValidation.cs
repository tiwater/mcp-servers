using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public sealed record OpenXmlValidationIssue(
    string Description,
    string? Path,
    string Id,
    string? CompatibilityCode = null);

public sealed record OpenXmlValidationResult(
    string Schema,
    bool Pass,
    string File,
    int ErrorCount,
    IReadOnlyList<OpenXmlValidationIssue> Errors,
    int WarningCount,
    IReadOnlyList<OpenXmlValidationIssue> Warnings);

public static class OpenXmlValidation
{
    private const string WordprocessingNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string UnexpectedUiPriorityDescription =
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:uiPriority'.";
    private const string TrailingUiPriorityCompatibilityCode =
        "wordprocessing-style-trailing-ui-priority";
    private const string StyleTableLayoutCompatibilityCode =
        "wordprocessing-style-table-layout";

    private static readonly HashSet<string> AllowedLeadingStyleMetadata =
    [
        "name",
        "aliases",
        "basedOn",
        "next",
        "link",
        "autoRedefine",
        "hidden",
        "semiHidden",
        "unhideWhenUsed",
        "qFormat",
    ];

    private static readonly HashSet<string> AllowedTrailingStylePayload =
    [
        "pPr",
        "rPr",
        "tblPr",
        "trPr",
        "tcPr",
        "tblStylePr",
    ];

    public static int Run(string[] args)
    {
        if (args.Length < 1) throw new InvalidOperationException("validate-openxml requires <input.docx>");
        var result = Validate(args[0]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    public static OpenXmlValidationResult Validate(string inputPath)
    {
        var path = Path.GetFullPath(inputPath);
        using var document = WordprocessingDocument.Open(path, false);
        var errors = new List<OpenXmlValidationIssue>();
        var warnings = new List<OpenXmlValidationIssue>();

        foreach (var error in new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document))
        {
            var issue = new OpenXmlValidationIssue(
                error.Description,
                error.Path?.XPath,
                error.Id);
            if (IsTrailingUiPriorityCompatibilityWarning(error))
            {
                warnings.Add(issue with { CompatibilityCode = TrailingUiPriorityCompatibilityCode });
            }
            else if (IsStyleTableLayoutCompatibilityWarning(error))
            {
                warnings.Add(issue with { CompatibilityCode = StyleTableLayoutCompatibilityCode });
            }
            else
            {
                errors.Add(issue);
            }
        }

        return new OpenXmlValidationResult(
            "tiwater.docx.openxml-validation/v1",
            errors.Count == 0,
            path,
            errors.Count,
            errors.Take(100).ToList(),
            warnings.Count,
            warnings.Take(100).ToList());
    }

    private static bool IsTrailingUiPriorityCompatibilityWarning(ValidationErrorInfo error)
    {
        if (error.Id != "Sch_UnexpectedElementContentExpectingComplex"
            || error.Description != UnexpectedUiPriorityDescription
            || error.Part?.Uri.OriginalString != "/word/styles.xml"
            || error.Node is not Style style
            || error.RelatedNode is not UIPriority uiPriority
            || style.Parent is not Styles styles
            || !ReferenceEquals(uiPriority.Parent, style))
        {
            return false;
        }

        var children = style.ChildElements.ToList();
        var uiPriorityIndex = children.IndexOf(uiPriority);
        if (uiPriorityIndex <= 0
            || children.Count(child => child is UIPriority) != 1
            || children.Take(uiPriorityIndex).Any(child =>
                child.NamespaceUri != WordprocessingNamespace
                || !AllowedLeadingStyleMetadata.Contains(child.LocalName))
            || children.Skip(uiPriorityIndex + 1).Any(child =>
                child.NamespaceUri != WordprocessingNamespace
                || !AllowedTrailingStylePayload.Contains(child.LocalName)))
        {
            return false;
        }

        var normalizedStyle = NormalizeStyleForCompatibilityProof(style, removeTableLayout: true);
        return IsValidIsolatedStyle(normalizedStyle, styles);
    }

    private static bool IsStyleTableLayoutCompatibilityWarning(ValidationErrorInfo error)
    {
        if (error.Id != "Sch_InvalidElementContentExpectingComplex"
            || error.Part?.Uri.OriginalString != "/word/styles.xml"
            || error.RelatedNode is not OpenXmlUnknownElement tableLayout
            || tableLayout.LocalName != "tblLayout"
            || tableLayout.NamespaceUri != WordprocessingNamespace
            || tableLayout.Parent is not StyleTableProperties tableProperties
            || tableProperties.Parent is not Style style
            || style.Parent is not Styles styles
            || tableProperties.ChildElements.Count(child =>
                child.LocalName == "tblLayout"
                && child.NamespaceUri == WordprocessingNamespace) != 1)
        {
            return false;
        }

        var normalizedStyle = NormalizeStyleForCompatibilityProof(style, removeTableLayout: true);
        return IsValidIsolatedStyle(normalizedStyle, styles);
    }

    private static Style NormalizeStyleForCompatibilityProof(
        Style style,
        bool removeTableLayout)
    {
        var normalizedStyle = (Style)style.CloneNode(true);
        if (removeTableLayout)
        {
            foreach (var tableLayout in normalizedStyle
                .Elements<StyleTableProperties>()
                .SelectMany(properties => properties.ChildElements)
                .Where(child =>
                    child.LocalName == "tblLayout"
                    && child.NamespaceUri == WordprocessingNamespace)
                .ToList())
            {
                tableLayout.Remove();
            }
        }

        var normalizedUiPriority = normalizedStyle.GetFirstChild<UIPriority>();
        var canonicalInsertionPoint = normalizedStyle.ChildElements.FirstOrDefault(child =>
            child.LocalName is "semiHidden" or "unhideWhenUsed" or "qFormat");
        if (normalizedUiPriority is not null && canonicalInsertionPoint is not null)
        {
            normalizedUiPriority.Remove();
            normalizedStyle.InsertBefore(normalizedUiPriority, canonicalInsertionPoint);
        }

        return normalizedStyle;
    }

    private static bool IsValidIsolatedStyle(Style normalizedStyle, Styles styles)
    {
        var normalizedStyles = (Styles)styles.CloneNode(false);
        normalizedStyles.Append(normalizedStyle);
        return !new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(normalizedStyles)
            .Any();
    }
}
