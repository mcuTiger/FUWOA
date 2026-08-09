using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fuwoa.AddIn.Commands
{
    public class SplitByColumnCommand
    {
        public void Execute(Excel.Application app)
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
                    int splitCol = selection.Column;

                    int lastRow = sourceSheet.Cells[sourceSheet.Rows.Count, splitCol]
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

                    bool isFiltered = sourceSheet.AutoFilterMode &&
                                      sourceSheet.AutoFilter != null &&
                                      sourceSheet.AutoFilter.FilterMode;

                    int lastCol = sourceSheet.Cells[headerRow, sourceSheet.Columns.Count]
                        .End[Excel.XlDirection.xlToLeft].Column;

                    int dataStartRow = headerRow + 1;

                    // 筛选态：先收集可见行号（纯整数运算，避免逐格 COM）
                    HashSet<int> visibleRows = null;
                    if (isFiltered)
                    {
                        visibleRows = new HashSet<int>();
                        var splitRange = sourceSheet.Range[
                            sourceSheet.Cells[dataStartRow, splitCol],
                            sourceSheet.Cells[lastRow, splitCol]];
                        Excel.Range visible;
                        try
                        {
                            visible = splitRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                        }
                        catch
                        {
                            MessageBox.Show(LanguageManager.Get("noDataBelow"), "FUWOA",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        foreach (Excel.Range area in visible.Areas)
                            for (int r = area.Row; r < area.Row + area.Rows.Count; r++)
                                visibleRows.Add(r);
                    }

                    // 分块批量读取 + 内存分组
                    const int ChunkRows = 20000;
                    var groups = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
                    var order = new List<string>();

                    for (int start = dataStartRow; start <= lastRow; start += ChunkRows)
                    {
                        int end = Math.Min(lastRow, start + ChunkRows - 1);
                        var rng = sourceSheet.Range[
                            sourceSheet.Cells[start, 1],
                            sourceSheet.Cells[end, lastCol]];
                        object raw = rng.Value2;
                        object[,] data;

                        if (raw is object[,] arr)
                        {
                            data = arr;
                        }
                        else
                        {
                            data = new object[1, lastCol];
                            for (int c = 1; c <= lastCol; c++)
                                data[0, c - 1] = raw;
                        }

                        int n = end - start + 1;
                        for (int r = 1; r <= n; r++)
                        {
                            int srcRow = start + r - 1;
                            if (visibleRows != null && !visibleRows.Contains(srcRow))
                                continue;

                            object keyObj = data[r, splitCol];
                            string key = keyObj?.ToString()?.Trim() ?? "";
                            if (!groups.TryGetValue(key, out var list))
                            {
                                list = new List<object[]>();
                                groups[key] = list;
                                order.Add(key);
                            }
                            var rowArr = new object[lastCol];
                            for (int c = 1; c <= lastCol; c++)
                                rowArr[c - 1] = data[r, c];
                            list.Add(rowArr);
                        }
                    }

                    if (groups.Count == 0)
                    {
                        MessageBox.Show(
                            LanguageManager.Get("noDataBelow"),
                            "FUWOA",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    // 收集已有工作表名
                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Excel.Worksheet ws in app.Worksheets)
                        existingNames.Add(ws.Name);

                    // 读取标题行
                    var headers = new object[lastCol];
                    for (int c = 1; c <= lastCol; c++)
                        headers[c - 1] = sourceSheet.Cells[headerRow, c].Value2;

                    Excel.Worksheet anchor = sourceSheet;
                    Excel.Worksheet firstNew = null;
                    int total = order.Count;
                    int done = 0;

                    foreach (string key in order)
                    {
                        done++;
                        app.StatusBar = $"FUWOA – {LanguageManager.Get("splitByColumn")}: {key} ({done}/{total})";

                        string finalName = MakeUniqueName(existingNames, SanitizeSheetName(key));
                        existingNames.Add(finalName);

                        var newSheet = (Excel.Worksheet)app.Worksheets.Add(
                            Type.Missing, anchor, Type.Missing, Type.Missing);
                        anchor = newSheet;
                        firstNew = firstNew ?? newSheet;
                        newSheet.Name = finalName;

                        var rows = groups[key];
                        var block = new object[rows.Count + 1, lastCol];
                        for (int c = 1; c <= lastCol; c++)
                            block[0, c - 1] = headers[c - 1];
                        for (int i = 0; i < rows.Count; i++)
                        {
                            var src = rows[i];
                            for (int c = 1; c <= lastCol; c++)
                                block[i + 1, c - 1] = src[c - 1];
                        }

                        var target = newSheet.Range[
                            newSheet.Cells[1, 1],
                            newSheet.Cells[rows.Count + 1, lastCol]];
                        target.Value2 = block;

                        newSheet.Range[
                            newSheet.Cells[1, 1],
                            newSheet.Cells[1, lastCol]].Font.Bold = true;

                        string lastColLetter = ColumnIndexToLetter(lastCol);
                        newSheet.Columns["A:" + lastColLetter].AutoFit();
                    }

                    firstNew?.Activate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"{LanguageManager.Get("splitFailed")}：{ex.Message}",
                    "FUWOA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                name = LanguageManager.Get("emptyCategory");
            name = name
                .Replace("[", "").Replace("]", "").Replace("/", "-")
                .Replace("*", "").Replace("?", "").Replace(":", "：")
                .Replace("\\", "-");
            if (name.Length > 31) name = name.Substring(0, 31);
            return name;
        }

        private static string MakeUniqueName(HashSet<string> existing, string baseName)
        {
            if (!existing.Contains(baseName))
                return baseName;

            for (int n = 2; n <= 999; n++)
            {
                string suffix = string.Format(" ({0})", n);
                int maxBaseLen = 31 - suffix.Length;
                string candidate = (baseName.Length > maxBaseLen
                    ? baseName.Substring(0, maxBaseLen)
                    : baseName) + suffix;
                if (!existing.Contains(candidate))
                    return candidate;
            }
            return baseName + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);
        }

        private static string ColumnIndexToLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }
    }
}
