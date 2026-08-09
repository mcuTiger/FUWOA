using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Fuwoa.Core.ExportCount;

namespace Fuwoa.AddIn.Commands
{
    /// <summary>
    /// 导出计数命令。
    /// </summary>
    public class ExportCountCommand
    {
        public void Execute(Excel.Application app, SortMode sortMode = SortMode.ByCount,
            bool descending = true, bool showPercentage = false)
        {
            try
            {
                if (app == null)
                {
                    MessageBox.Show(
                        LanguageManager.Get("noExcelApp"),
                        "FUWOA",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null || selection.Cells.Count != 1)
                {
                    MessageBox.Show(
                        LanguageManager.Get("selectOneCell"),
                        "FUWOA",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                using (new ExcelGuard(app))
                {
                    Excel.Worksheet sourceSheet = selection.Worksheet;
                    int headerRow = selection.Row;
                    int columnIndex = selection.Column;

                    string headerText = sourceSheet.Cells[headerRow, columnIndex].Text?.ToString();
                    if (string.IsNullOrWhiteSpace(headerText))
                    {
                        headerText = LanguageManager.Get("column") + columnIndex;
                    }

                    int lastRow = sourceSheet.Cells[sourceSheet.Rows.Count, columnIndex]
                        .End[Excel.XlDirection.xlUp].Row;

                    if (lastRow <= headerRow)
                    {
                        MessageBox.Show(
                            LanguageManager.Get("noDataBelow"),
                            "FUWOA",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    int dataStartRow = headerRow + 1;

                    bool isFiltered = sourceSheet.AutoFilterMode &&
                                      sourceSheet.AutoFilter != null &&
                                      sourceSheet.AutoFilter.FilterMode;

                    var dataRange = sourceSheet.Range[
                        sourceSheet.Cells[dataStartRow, columnIndex],
                        sourceSheet.Cells[lastRow, columnIndex]];

                    var items = new System.Collections.Generic.List<string>();

                    if (isFiltered)
                    {
                        Excel.Range visible;
                        try
                        {
                            visible = dataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                        }
                        catch
                        {
                            MessageBox.Show(LanguageManager.Get("noDataBelow"), "FUWOA",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        foreach (Excel.Range area in visible.Areas)
                        {
                            object raw = area.Value2;
                            if (raw is object[,] arr)
                            {
                                for (int i = 1; i <= arr.GetLength(0); i++)
                                    items.Add(arr[i, 1]?.ToString() ?? "");
                            }
                            else
                            {
                                items.Add(raw?.ToString() ?? "");
                            }
                        }
                    }
                    else
                    {
                        object raw = dataRange.Value2;
                        if (raw is object[,] arr)
                        {
                            for (int i = 1; i <= arr.GetLength(0); i++)
                                items.Add(arr[i, 1]?.ToString() ?? "");
                        }
                        else
                        {
                            items.Add(raw?.ToString() ?? "");
                        }
                    }

                    var service = new ExportCountService();
                    var result = service.Compute(items.ToArray(), sortMode, descending);

                    Excel.Worksheet newSheet = (Excel.Worksheet)app.Worksheets.Add(
                        Type.Missing, sourceSheet, Type.Missing, Type.Missing);

                    string sheetName = headerText + LanguageManager.Get("count");
                    sheetName = sheetName
                        .Replace("[", "").Replace("]", "").Replace("/", "-")
                        .Replace("*", "").Replace("?", "").Replace(":", "：")
                        .Replace("\\", "-");
                    if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);

                    string finalName = sheetName;
                    var existingNames = new System.Collections.Generic.HashSet<string>(
                        StringComparer.OrdinalIgnoreCase);
                    foreach (Excel.Worksheet ws in app.Worksheets)
                        existingNames.Add(ws.Name);

                    if (existingNames.Contains(sheetName))
                    {
                        for (int n = 2; n <= 999; n++)
                        {
                            string suffix = string.Format(" ({0})", n);
                            int suffixLen = suffix.Length;
                            string candidate = sheetName;
                            if (candidate.Length + suffixLen > 31)
                                candidate = candidate.Substring(0, 31 - suffixLen);
                            candidate += suffix;
                            if (!existingNames.Contains(candidate))
                            {
                                finalName = candidate;
                                break;
                            }
                        }
                    }
                    newSheet.Name = finalName;

                    newSheet.Range["A:A"].NumberFormat = "@";

                    int cols = showPercentage ? 3 : 2;
                    var block = new object[result.Count + 1, cols];
                    block[0, 0] = headerText;
                    block[0, 1] = LanguageManager.Get("count");
                    if (showPercentage)
                        block[0, 2] = LanguageManager.Get("percentageColumn");

                    int totalCount = 0;
                    for (int i = 0; i < result.Count; i++)
                    {
                        block[i + 1, 0] = result[i].Value;
                        block[i + 1, 1] = result[i].Count;
                        totalCount += result[i].Count;
                    }
                    if (showPercentage && totalCount > 0)
                    {
                        newSheet.Columns["C"].NumberFormat = "0.0%";
                        newSheet.Columns["C"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        for (int i = 0; i < result.Count; i++)
                            block[i + 1, 2] = (double)result[i].Count / totalCount;
                    }

                    var target = newSheet.Range[
                        newSheet.Cells[1, 1],
                        newSheet.Cells[result.Count + 1, cols]];
                    target.Value2 = block;

                    newSheet.Range["A1:B1"].Font.Bold = true;
                    if (showPercentage) newSheet.Cells[1, 3].Font.Bold = true;
                    newSheet.Columns[showPercentage ? "A:C" : "A:B"].AutoFit();
                    newSheet.Activate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"{LanguageManager.Get("exportFailed")}：{ex.Message}",
                    "FUWOA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
