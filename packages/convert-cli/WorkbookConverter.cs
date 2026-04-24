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
        var sourceWorkbook = new HSSFWorkbook(inputStream);
        using var targetWorkbook = new XSSFWorkbook();

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
