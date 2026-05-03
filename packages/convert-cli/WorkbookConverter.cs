using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Dockit.Convert;

public static class WorkbookConverter
{
    public static void ConvertXlsToXlsx(string input, string output)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

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
