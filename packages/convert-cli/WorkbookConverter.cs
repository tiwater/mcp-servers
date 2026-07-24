using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Security.Cryptography;
using System.Text.Json;

namespace Dockit.Convert;

public sealed class AuthoritativeSpreadsheetRuntimeException : Exception
{
    public AuthoritativeSpreadsheetRuntimeException(string message, Exception? innerException = null)
        : base(message, innerException) { }
}

public static class WorkbookConverter
{
    private const string InspectionCacheSchema = "tiwater.xls-to-xlsx-cache/v1";

    public sealed record ConversionResult(string Backend, string? FallbackReason = null);
    private sealed record InspectionCacheManifest(string Schema, string InputSha256, string OutputSha256, string Backend);

    public static ConversionResult ConvertXlsToXlsx(string input, string output)
        => ConvertXlsToXlsx(input, output, requireWpsForAuthority: false);

    public static ConversionResult ConvertXlsToXlsxForInspection(string input, string output)
    {
        input = Path.GetFullPath(input);
        output = Path.GetFullPath(output);
        if (!File.Exists(input)) throw new InvalidOperationException($"Input file not found: {input}");

        var inputSha256 = FileSha256(input);
        var cacheDirectory = Path.Combine(InspectionCacheRoot(), "xls-to-xlsx", "wps-spreadsheet", "v1", inputSha256);
        var cachedWorkbook = Path.Combine(cacheDirectory, "normalized.xlsx");
        var manifestPath = Path.Combine(cacheDirectory, "manifest.json");
        Directory.CreateDirectory(cacheDirectory);

        using (WpsRpcSession.AcquireContentLease(Path.Combine(cacheDirectory, "conversion.lock"), TimeSpan.FromMinutes(5)))
        {
            if (!ValidInspectionCache(cachedWorkbook, manifestPath, inputSha256))
            {
                File.Delete(cachedWorkbook);
                File.Delete(manifestPath);
                var candidate = Path.Combine(cacheDirectory, $"normalized-{Guid.NewGuid():N}.xlsx");
                try
                {
                    var conversion = ConvertXlsToXlsx(input, candidate, requireWpsForAuthority: true);
                    if (!string.Equals(conversion.Backend, "wps-spreadsheet", StringComparison.Ordinal)
                        || !string.IsNullOrWhiteSpace(conversion.FallbackReason))
                        throw new InvalidOperationException("Authoritative XLS inspection cache requires WPS Spreadsheet conversion without fallback.");
                    ValidateXlsxPackage(candidate);
                    var outputSha256 = FileSha256(candidate);
                    File.Move(candidate, cachedWorkbook, overwrite: true);
                    AtomicWrite(
                        manifestPath,
                        JsonSerializer.Serialize(new InspectionCacheManifest(
                            InspectionCacheSchema,
                            inputSha256,
                            outputSha256,
                            conversion.Backend)));
                }
                finally
                {
                    File.Delete(candidate);
                }
            }
        }

        var outputDirectory = Path.GetDirectoryName(output);
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        File.Copy(cachedWorkbook, output, overwrite: true);
        ValidateXlsxPackage(output);
        return new ConversionResult("wps-spreadsheet");
    }

    private static bool ValidInspectionCache(string workbookPath, string manifestPath, string inputSha256)
    {
        try
        {
            if (!File.Exists(workbookPath) || !File.Exists(manifestPath)) return false;
            var manifest = JsonSerializer.Deserialize<InspectionCacheManifest>(File.ReadAllText(manifestPath));
            if (manifest is null
                || !string.Equals(manifest.Schema, InspectionCacheSchema, StringComparison.Ordinal)
                || !string.Equals(manifest.InputSha256, inputSha256, StringComparison.Ordinal)
                || !string.Equals(manifest.Backend, "wps-spreadsheet", StringComparison.Ordinal)
                || !string.Equals(manifest.OutputSha256, FileSha256(workbookPath), StringComparison.Ordinal))
                return false;
            ValidateXlsxPackage(workbookPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string InspectionCacheRoot()
    {
        var configured = Environment.GetEnvironmentVariable("TIWATER_OFFICE_CACHE_ROOT")?.Trim();
        if (!string.IsNullOrWhiteSpace(configured)) return Path.GetFullPath(configured);
        var xdg = Environment.GetEnvironmentVariable("XDG_CACHE_HOME")?.Trim();
        if (!string.IsNullOrWhiteSpace(xdg)) return Path.Combine(Path.GetFullPath(xdg), "tiwater", "office-conversions");
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".cache", "tiwater", "office-conversions");
    }

    private static string FileSha256(string path)
    {
        using var stream = File.OpenRead(path);
        return System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static void ValidateXlsxPackage(string path)
    {
        using var stream = File.OpenRead(path);
        using var workbook = new XSSFWorkbook(stream);
        if (workbook.NumberOfSheets < 1) throw new InvalidOperationException($"WPS RPC produced an XLSX without worksheets: {path}");
    }

    private static void AtomicWrite(string path, string content)
    {
        var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(temporary, content);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            File.Delete(temporary);
        }
    }

    private static ConversionResult ConvertXlsToXlsx(string input, string output, bool requireWpsForAuthority)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var requiredBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_XLSX_BACKEND")?.Trim();
        var requireWps = requireWpsForAuthority || string.Equals(requiredBackend, "wps-spreadsheet", StringComparison.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(requiredBackend) && !requireWps)
        {
            throw new InvalidOperationException($"Unsupported required XLSX backend: {requiredBackend}");
        }

        if (WpsSpreadsheetConverter.IsAvailable())
        {
            try
            {
                WpsSpreadsheetConverter.ConvertXlsToXlsx(input, output);
                return new ConversionResult("wps-spreadsheet");
            }
            catch (Exception ex)
            {
                if (requireWps)
                {
                    throw new AuthoritativeSpreadsheetRuntimeException($"Required WPS Spreadsheet XLS conversion failed: {ex.Message}", ex);
                }
                var fallbackReason = $"WPS RPC conversion failed: {ex.Message}";
                var fallback = ConvertXlsToXlsxWithoutWps(input, output);
                return fallback with
                {
                    FallbackReason = string.IsNullOrWhiteSpace(fallback.FallbackReason)
                        ? fallbackReason
                        : $"{fallbackReason}; {fallback.FallbackReason}",
                };
            }
        }

        if (LimaWpsWriterPdfConverter.IsAvailable())
        {
            try
            {
                LimaWpsWriterPdfConverter.ConvertSpreadsheetToXlsx(input, output);
                return new ConversionResult("wps-spreadsheet");
            }
            catch (Exception ex)
            {
                if (requireWps)
                {
                    throw new AuthoritativeSpreadsheetRuntimeException($"Required Lima WPS Spreadsheet XLS conversion failed: {ex.Message}", ex);
                }
                var fallbackReason = $"Lima WPS RPC conversion failed: {ex.Message}";
                var fallback = ConvertXlsToXlsxWithoutWps(input, output);
                return fallback with
                {
                    FallbackReason = string.IsNullOrWhiteSpace(fallback.FallbackReason)
                        ? fallbackReason
                        : $"{fallbackReason}; {fallback.FallbackReason}",
                };
            }
        }

        if (requireWps)
        {
            throw new AuthoritativeSpreadsheetRuntimeException(
                "WPS Spreadsheet XLS conversion was required but neither local WPS RPC nor a configured Lima WPS runtime is available.");
        }

        return ConvertXlsToXlsxWithoutWps(input, output);
    }

    private static ConversionResult ConvertXlsToXlsxWithoutWps(string input, string output)
    {
        var soffice = OfficeConverter.FindSofficeBinary();
        if (!string.IsNullOrWhiteSpace(soffice))
        {
            try
            {
                OfficeConverter.ConvertXlsToXlsx(input, output, soffice);
                return new ConversionResult("libreoffice");
            }
            catch (Exception ex)
            {
                var fallbackReason = $"LibreOffice/soffice conversion failed: {ex.Message}";
                ConvertXlsToXlsxWithNpoi(input, output);
                return new ConversionResult("npoi", fallbackReason);
            }
        }

        ConvertXlsToXlsxWithNpoi(input, output);
        return new ConversionResult("npoi", "WPS RPC and LibreOffice/soffice not found");
    }

    private static void ConvertXlsToXlsxWithNpoi(string input, string output)
    {
        using var inputStream = File.OpenRead(input);
        HSSFWorkbook sourceWorkbook;
        try
        {
            sourceWorkbook = new HSSFWorkbook(inputStream);
        }
        catch (Exception ex)
        {
            throw ClassifyOpenWorkbookError(input, ex);
        }
        using var targetWorkbook = new XSSFWorkbook();
        var styleMap = new Dictionary<short, ICellStyle>();

        for (var sheetIndex = 0; sheetIndex < sourceWorkbook.NumberOfSheets; sheetIndex++)
        {
            var sourceSheet = sourceWorkbook.GetSheetAt(sheetIndex);
            var targetSheet = targetWorkbook.CreateSheet(sourceSheet.SheetName);

            for (var rowIndex = sourceSheet.FirstRowNum; rowIndex <= sourceSheet.LastRowNum; rowIndex++)
            {
                var sourceRow = sourceSheet.GetRow(rowIndex);
                if (sourceRow is null)
                {
                    continue;
                }

                var targetRow = targetSheet.CreateRow(rowIndex);
                targetRow.Height = sourceRow.Height;

                for (var cellIndex = 0; cellIndex < sourceRow.LastCellNum; cellIndex++)
                {
                    var sourceCell = sourceRow.GetCell(cellIndex);
                    if (sourceCell is null)
                    {
                        continue;
                    }

                    var targetCell = targetRow.CreateCell(cellIndex);
                    CopyCellValue(sourceCell, targetCell);
                    CopyCellStyle(targetWorkbook, sourceCell, targetCell, styleMap);
                }
            }

            for (var i = 0; i < sourceSheet.NumMergedRegions; i++)
            {
                targetSheet.AddMergedRegion(sourceSheet.GetMergedRegion(i));
            }

            for (var i = 0; i <= 255; i++)
            {
                try
                {
                    targetSheet.SetColumnWidth(i, sourceSheet.GetColumnWidth(i));
                }
                catch
                {
                    // ignore sparse column width issues
                }
            }
        }

        var outputDir = Path.GetDirectoryName(output);
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        using var outputStream = File.Create(output);
        targetWorkbook.Write(outputStream);
    }

    public static Exception ClassifyOpenWorkbookError(string input, Exception ex)
    {
        var message = ex.Message ?? string.Empty;
        var normalized = message.ToLowerInvariant();
        if (normalized.Contains("password") || normalized.Contains("encrypted") || normalized.Contains("poi 4.2"))
        {
            return new InvalidOperationException(
                $"Encrypted or password-protected XLS is not supported for conversion: {input}",
                ex);
        }

        return new InvalidOperationException(
            $"Failed to open legacy XLS for conversion: {input} :: {message}",
            ex);
    }

    private static void CopyCellStyle(
        XSSFWorkbook targetWorkbook,
        ICell sourceCell,
        ICell targetCell,
        Dictionary<short, ICellStyle> styleMap)
    {
        var sourceStyle = sourceCell.CellStyle;
        if (sourceStyle is null)
        {
            return;
        }

        if (!styleMap.TryGetValue(sourceStyle.Index, out var targetStyle))
        {
            targetStyle = targetWorkbook.CreateCellStyle();
            CopyStyleProperties(sourceStyle, targetStyle);
            styleMap[sourceStyle.Index] = targetStyle;
        }

        targetCell.CellStyle = targetStyle;
    }

    private static void CopyStyleProperties(ICellStyle sourceStyle, ICellStyle targetStyle)
    {
        targetStyle.Alignment = sourceStyle.Alignment;
        targetStyle.VerticalAlignment = sourceStyle.VerticalAlignment;
        targetStyle.BorderBottom = sourceStyle.BorderBottom;
        targetStyle.BorderLeft = sourceStyle.BorderLeft;
        targetStyle.BorderRight = sourceStyle.BorderRight;
        targetStyle.BorderTop = sourceStyle.BorderTop;
        targetStyle.BottomBorderColor = sourceStyle.BottomBorderColor;
        targetStyle.LeftBorderColor = sourceStyle.LeftBorderColor;
        targetStyle.RightBorderColor = sourceStyle.RightBorderColor;
        targetStyle.TopBorderColor = sourceStyle.TopBorderColor;
        targetStyle.DataFormat = sourceStyle.DataFormat;
        targetStyle.FillBackgroundColor = sourceStyle.FillBackgroundColor;
        targetStyle.FillForegroundColor = sourceStyle.FillForegroundColor;
        targetStyle.FillPattern = sourceStyle.FillPattern;
        targetStyle.Indention = sourceStyle.Indention;
        targetStyle.IsLocked = sourceStyle.IsLocked;
        targetStyle.Rotation = sourceStyle.Rotation;
        targetStyle.ShrinkToFit = sourceStyle.ShrinkToFit;
        targetStyle.WrapText = sourceStyle.WrapText;
    }

    private static void CopyCellValue(ICell sourceCell, ICell targetCell)
    {
        switch (sourceCell.CellType)
        {
            case CellType.String:
                targetCell.SetCellValue(sourceCell.StringCellValue);
                break;
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(sourceCell))
                {
                    var dateValue = sourceCell.DateCellValue;
                    if (dateValue.HasValue)
                    {
                        targetCell.SetCellValue(dateValue.Value);
                    }
                    else
                    {
                        targetCell.SetBlank();
                    }
                }
                else
                {
                    targetCell.SetCellValue(sourceCell.NumericCellValue);
                }
                break;
            case CellType.Boolean:
                targetCell.SetCellValue(sourceCell.BooleanCellValue);
                break;
            case CellType.Formula:
                targetCell.SetCellFormula(sourceCell.CellFormula);
                break;
            case CellType.Blank:
                targetCell.SetBlank();
                break;
            default:
                targetCell.SetCellValue(sourceCell.ToString());
                break;
        }
    }
}
