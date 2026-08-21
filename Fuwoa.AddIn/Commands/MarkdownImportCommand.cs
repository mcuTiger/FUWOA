using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fuwoa.AddIn.Commands
{
    /// <summary>Markdown 表格数据来源。</summary>
    public enum MarkdownImportSource
    {
        File,
        Clipboard
    }

    /// <summary>解析出的一个 Markdown 表格。</summary>
    public class MarkdownTable
    {
        public List<string[]> Rows { get; } = new List<string[]>();
        public int Columns { get; set; }

        public void AddRow(string[] cells)
        {
            Rows.Add(cells);
            if (cells.Length > Columns) Columns = cells.Length;
        }

        public string Preview
        {
            get
            {
                if (Rows.Count == 0) return "";
                var first = Rows[0];
                var sb = new StringBuilder();
                int shown = Math.Min(3, first.Length);
                for (int i = 0; i < shown; i++)
                {
                    if (i > 0) sb.Append(" | ");
                    sb.Append(first[i]);
                }
                string s = sb.ToString();
                return s.Length > 40 ? s.Substring(0, 40) + "…" : s;
            }
        }
    }

    /// <summary>
    /// Markdown 表格导入命令：从 .md 文件或剪贴板解析 Markdown 表格，
    /// 弹出选择框勾选后，每个表格写入一个新建工作表。
    /// </summary>
    public class MarkdownImportCommand
    {
        public void Execute(Excel.Application app, MarkdownImportSource source)
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

                string text = null;
                if (source == MarkdownImportSource.File)
                {
                    using (var dlg = new OpenFileDialog())
                    {
                        dlg.Filter = "Markdown 文件 (*.md;*.markdown;*.txt)|*.md;*.markdown;*.txt|所有文件 (*.*)|*.*";
                        dlg.Title = LanguageManager.Get("mdImportFile");
                        if (dlg.ShowDialog() != DialogResult.OK) return;
                        text = ReadFileWithFallback(dlg.FileName);
                    }
                }
                else
                {
                    text = Clipboard.ContainsText() ? Clipboard.GetText() : null;
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        MessageBox.Show(
                            LanguageManager.Get("mdImportClipboardEmpty"),
                            "FUWOA",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }
                }

                var tables = MarkdownTableParser.Parse(text);
                if (tables.Count == 0)
                {
                    MessageBox.Show(
                        LanguageManager.Get("mdImportNoTable"),
                        "FUWOA",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                int[] selected = ShowTablePicker(tables);
                if (selected == null || selected.Length == 0) return;

                using (new ExcelGuard(app))
                {
                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Excel.Worksheet ws in app.Worksheets)
                        existingNames.Add(ws.Name);

                    Excel.Worksheet anchor = app.ActiveSheet as Excel.Worksheet
                        ?? (Excel.Worksheet)app.Worksheets[1];

                    int tableNo = 1;
                    foreach (int idx in selected)
                    {
                        var table = tables[idx];

                        string baseName = MakeBaseName(table, tableNo);
                        string finalName = MakeUniqueName(existingNames, SanitizeSheetName(baseName));
                        existingNames.Add(finalName);

                        var newSheet = (Excel.Worksheet)app.Worksheets.Add(
                            Type.Missing, anchor, Type.Missing, Type.Missing);
                        anchor = newSheet;
                        newSheet.Name = finalName;

                        var block = new object[table.Rows.Count, table.Columns];
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            var src = table.Rows[r];
                            for (int c = 0; c < table.Columns; c++)
                                block[r, c] = c < src.Length ? src[c] : "";
                        }

                        var target = newSheet.Range[
                            newSheet.Cells[1, 1],
                            newSheet.Cells[table.Rows.Count, table.Columns]];
                        target.Value2 = block;

                        newSheet.Range[
                            newSheet.Cells[1, 1],
                            newSheet.Cells[1, table.Columns]].Font.Bold = true;

                        string lastColLetter = ColumnIndexToLetter(table.Columns);
                        newSheet.Columns["A:" + lastColLetter].AutoFit();

                        tableNo++;
                    }

                    app.StatusBar = string.Format(LanguageManager.Get("mdImportDone"), selected.Length);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"{LanguageManager.Get("mdImportFailed")}：{ex.Message}",
                    "FUWOA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>读取文件，UTF-8 优先，解码失败回退 GB18030。</summary>
        private static string ReadFileWithFallback(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            try
            {
                var utf8 = new UTF8Encoding(false, true);
                return utf8.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return Encoding.GetEncoding("GB18030").GetString(bytes);
            }
        }

        /// <summary>弹出多选表格对话框，返回选中的表格索引；取消返回 null。</summary>
        private static int[] ShowTablePicker(List<MarkdownTable> tables)
        {
            using (var form = new Form())
            {
                form.Text = LanguageManager.Get("mdImportSelect");
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Width = 560;
                form.Height = 440;
                form.MinimizeBox = false;
                form.MaximizeBox = false;

                var list = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    CheckOnClick = true
                };
                for (int i = 0; i < tables.Count; i++)
                {
                    var t = tables[i];
                    string header = string.IsNullOrEmpty(t.Preview)
                        ? "（空表头）"
                        : t.Preview;
                    list.Items.Add(
                        $"[{i + 1}] {header}  ——  {t.Rows.Count} 行 × {t.Columns} 列",
                        true);
                }

                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                var bottom = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 44,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(8, 8, 8, 4)
                };
                bottom.Controls.Add(cancel);
                bottom.Controls.Add(ok);

                form.Controls.Add(list);
                form.Controls.Add(bottom);
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                if (form.ShowDialog() != DialogResult.OK)
                    return null;

                var result = new List<int>();
                for (int i = 0; i < list.Items.Count; i++)
                    if (list.GetItemChecked(i))
                        result.Add(i);
                return result.ToArray();
            }
        }

        private static string MakeBaseName(MarkdownTable table, int tableNo)
        {
            if (table.Rows.Count > 0 && table.Rows[0].Length > 0)
            {
                string first = table.Rows[0][0]?.Trim() ?? "";
                if (!string.IsNullOrEmpty(first))
                    return first;
            }
            return LanguageManager.Get("mdImportTable") + tableNo;
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                name = LanguageManager.Get("mdImportTable");
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

    /// <summary>Markdown 表格解析器：只识别表格块，忽略 YAML 头、引用、代码块、列表等。</summary>
    public static class MarkdownTableParser
    {
        public static List<MarkdownTable> Parse(string text)
        {
            var tables = new List<MarkdownTable>();
            if (string.IsNullOrEmpty(text)) return tables;

            string[] lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            bool inCodeBlock = false;
            bool inYaml = false;
            MarkdownTable current = null;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                // 代码块：整块跳过
                if (line.StartsWith("```"))
                {
                    inCodeBlock = !inCodeBlock;
                    current = null;
                    continue;
                }
                if (inCodeBlock)
                {
                    current = null;
                    continue;
                }

                // YAML 头：`---` 开始，要求下一行形如 `Key:`；到下一个 `---` 结束
                if (inYaml)
                {
                    if (line == "---")
                        inYaml = false;
                    current = null;
                    continue;
                }
                if (line == "---" && i + 1 < lines.Length && IsYamlKeyLine(lines[i + 1].Trim()))
                {
                    inYaml = true;
                    current = null;
                    continue;
                }

                // 引用块：跳过
                if (line.StartsWith(">"))
                {
                    current = null;
                    continue;
                }

                // 表格行
                if (line.StartsWith("|"))
                {
                    string[] cells = SplitRow(line);

                    if (IsSeparatorRow(cells))
                    {
                        current = null;
                        continue;
                    }

                    if (current == null)
                    {
                        // 需确认下一行是分隔行才算表格
                        if (i + 1 < lines.Length)
                        {
                            string[] next = SplitRow(lines[i + 1].Trim());
                            if (IsSeparatorRow(next))
                            {
                                current = new MarkdownTable();
                                tables.Add(current);
                                current.AddRow(cells);
                                i++; // 跳过分隔行
                            }
                        }
                        continue;
                    }

                    current.AddRow(cells);
                    continue;
                }

                // 其他内容行：表格结束
                current = null;
            }

            return tables;
        }

        private static bool IsYamlKeyLine(string line)
        {
            if (string.IsNullOrEmpty(line)) return false;
            if (line == "---") return false;
            return Regex.IsMatch(line, @"^[A-Za-z_][A-Za-z0-9_]*:");
        }

        /// <summary>按 | 拆分一行，处理 \| 转义，去掉首尾空单元格。</summary>
        private static string[] SplitRow(string line)
        {
            string s = line.Trim();
            if (s.StartsWith("|")) s = s.Substring(1);
            if (s.EndsWith("|") && !s.EndsWith("\\|"))
                s = s.Substring(0, s.Length - 1);

            var parts = new List<string>();
            var cur = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\\' && i + 1 < s.Length && s[i + 1] == '|')
                {
                    cur.Append('|');
                    i++;
                    continue;
                }
                if (c == '|')
                {
                    parts.Add(cur.ToString().Trim());
                    cur.Clear();
                    continue;
                }
                cur.Append(c);
            }
            parts.Add(cur.ToString().Trim());
            return parts.ToArray();
        }

        /// <summary>分隔行：每个单元格为 `---`、`:---`、`---:`、`:---:` 或空。</summary>
        private static bool IsSeparatorRow(string[] cells)
        {
            if (cells == null || cells.Length == 0) return false;
            foreach (string cell in cells)
            {
                string t = cell.Trim();
                if (t.Length == 0) continue;
                if (!Regex.IsMatch(t, @"^:?-{2,}:?$")) return false;
            }
            return true;
        }
    }
}
