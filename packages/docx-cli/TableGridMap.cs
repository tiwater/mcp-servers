using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public class TableGridMap
{
    public int RowCount { get; }
    public int ColumnCount { get; }
    public TableCell?[,] Grid { get; }

    public TableGridMap(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        RowCount = rows.Count;

        // Determine grid width
        var gridWidth = 0;
        foreach (var row in rows)
        {
            var rowWidth = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                rowWidth += cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            }
            gridWidth = Math.Max(gridWidth, rowWidth);
        }
        ColumnCount = gridWidth;
        Grid = new TableCell?[RowCount, ColumnCount];

        // Track vertical merges per column index
        var activeVerticalMerges = new TableCell?[ColumnCount];

        for (var r = 0; r < RowCount; r++)
        {
            var row = rows[r];
            var cells = row.Elements<TableCell>().ToList();
            var cellIndex = 0;

            for (var col = 0; col < ColumnCount; col++)
            {
                // If this position is already populated by a vertical merge, skip it
                if (Grid[r, col] != null)
                {
                    continue;
                }

                if (cellIndex >= cells.Count)
                {
                    break;
                }

                var cell = cells[cellIndex];
                var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                var vMerge = cell.TableCellProperties?.VerticalMerge;

                if (vMerge != null)
                {
                    if (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Restart)
                    {
                        for (var i = 0; i < span; i++)
                        {
                            activeVerticalMerges[col + i] = cell;
                        }
                    }
                    else if (vMerge.Val.Value == MergedCellValues.Continue)
                    {
                        var mergedParent = activeVerticalMerges[col];
                        if (mergedParent != null)
                        {
                            cell = mergedParent;
                        }
                    }
                }
                else
                {
                    for (var i = 0; i < span; i++)
                    {
                        activeVerticalMerges[col + i] = null;
                    }
                }

                for (var i = 0; i < span; i++)
                {
                    Grid[r, col + i] = cell;
                }

                cellIndex++;
                col += span - 1;
            }
        }
    }

    public string GetRowContext(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= RowCount)
        {
            return string.Empty;
        }

        // Concatenate text from the first two cells in the row
        var cell1 = Grid[rowIndex, 0];
        var cell2 = ColumnCount > 1 ? Grid[rowIndex, 1] : null;

        var text1 = cell1 != null ? GetCellText(cell1) : string.Empty;
        var text2 = cell2 != null && !ReferenceEquals(cell1, cell2) ? GetCellText(cell2) : string.Empty;

        return $"{text1} {text2}".Trim();
    }

    public string GetColumnContext(int colIndex)
    {
        if (colIndex < 0 || colIndex >= ColumnCount)
        {
            return string.Empty;
        }

        // Concatenate text from the cells in row 0 and row 1 for this column (the headers)
        var cell0 = Grid[0, colIndex];
        var cell1 = RowCount > 1 ? Grid[1, colIndex] : null;
        var cell2 = RowCount > 2 ? Grid[2, colIndex] : null;

        var text0 = cell0 != null ? GetCellText(cell0) : string.Empty;
        var text1 = cell1 != null && !ReferenceEquals(cell0, cell1) ? GetCellText(cell1) : string.Empty;
        var text2 = cell2 != null && !ReferenceEquals(cell0, cell2) && !ReferenceEquals(cell1, cell2) ? GetCellText(cell2) : string.Empty;

        return $"{text0} {text1} {text2}".Trim();
    }

    private static string GetCellText(TableCell cell)
    {
        return string.Concat(cell.Descendants<Text>().Select(t => t.Text)).Trim();
    }
}
