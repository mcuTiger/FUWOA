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

                // 读取数据（用 Text 属性保留原始显示格式，避免 01-06 被截断）
                string[] items = new string[dataRowCount];
                for (int i = 0; i < dataRowCount; i++)
                {
                    items[i] = sourceSheet.Cells[dataStartRow + i, columnIndex].Text?.ToString()
                               ?? string.Empty;
                }

                // 调用核心服务
                var service = new ExportCountService();
                var result = service.Compute(items);

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
                newSheet.Name = sheetName;

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
