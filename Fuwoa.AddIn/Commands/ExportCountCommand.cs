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
            bool descending = true)
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

                Excel.Worksheet sourceSheet = selection.Worksheet;
                int headerRow = selection.Row;
                int columnIndex = selection.Column;

                // 读取真正的标题文本（选中的标题单元格）
                string headerText = sourceSheet.Cells[headerRow, columnIndex].Text?.ToString();
                if (string.IsNullOrWhiteSpace(headerText))
                {
                    headerText = LanguageManager.Get("column") + columnIndex;
                }

                // 读取该列从标题行下一行开始到末尾的所有数据
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
                int dataRowCount = lastRow - dataStartRow + 1;

                // 检测筛选状态
                bool isFiltered = sourceSheet.AutoFilterMode &&
                                  sourceSheet.AutoFilter != null &&
                                  sourceSheet.AutoFilter.FilterMode;

                // 读取数据（用 Text 属性保留原始显示格式，避免 01-06 被截断）
                var items = new System.Collections.Generic.List<string>();
                if (isFiltered)
                {
                    // 仅读取可见行
                    var dataRange = sourceSheet.Range[
                        sourceSheet.Cells[dataStartRow, columnIndex],
                        sourceSheet.Cells[lastRow, columnIndex]];
                    var visibleAreas = dataRange.SpecialCells(
                        Excel.XlCellType.xlCellTypeVisible);
                    foreach (Excel.Range area in visibleAreas.Areas)
                    {
                        foreach (Excel.Range cell in area.Cells)
                        {
                            items.Add(cell.Text?.ToString() ?? string.Empty);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        items.Add(sourceSheet.Cells[dataStartRow + i, columnIndex].Text?.ToString()
                                   ?? string.Empty);
                    }
                }

                // 调用核心服务
                var service = new ExportCountService();
                var result = service.Compute(items.ToArray(), sortMode, descending);

                // 创建新工作表
                Excel.Worksheet newSheet = (Excel.Worksheet)app.Worksheets.Add(
                    Type.Missing, sourceSheet, Type.Missing, Type.Missing);

                // 重命名为 "标题计数"（如 "管理编号计数"）
                string sheetName = headerText + LanguageManager.Get("count");
                // Excel 工作表名最长 31 字符，且不能含 [ ] / * ? : \
                sheetName = sheetName
                    .Replace("[", "").Replace("]", "").Replace("/", "-")
                    .Replace("*", "").Replace("?", "").Replace(":", "：")
                    .Replace("\\", "-");
                if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);

                // 重名时加序号后缀，避免覆盖用户已标注的旧表
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

                // 设置 A 列为文本格式，防止 Excel 自动格式化（如 01-06 → 1-6）
                newSheet.Range["A:A"].NumberFormat = "@";

                // 写入表头
                newSheet.Cells[1, 1].Value = headerText;
                newSheet.Cells[1, 2].Value = LanguageManager.Get("count");
                newSheet.Range["A1:B1"].Font.Bold = true;

                // 写入结果（A 列已设为文本格式，B 列为数字）
                for (int i = 0; i < result.Count; i++)
                {
                    newSheet.Cells[i + 2, 1].Value = result[i].Value;
                    newSheet.Cells[i + 2, 2].Value2 = result[i].Count;
                }

                // 自动调整列宽并激活
                newSheet.Columns["A:B"].AutoFit();
                newSheet.Activate();
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
