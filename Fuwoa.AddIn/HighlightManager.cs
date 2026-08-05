#if !RELEASE
using System;
using System.Reflection;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fuwoa.AddIn
{
    /// <summary>
    /// 行列高亮管理器：使用条件格式 + 命名范围实现高性能高亮。
    /// 切换选区只更新命名范围值，Excel 自动重算条件格式；
    /// 不逐格读写 Interior.Color，不破坏原始背景色。
    /// </summary>
    public class HighlightManager
    {
        private const string RegPath = @"SOFTWARE\FUWOA";
        private const string ColorKey = "HighlightColor";
        private const int DefaultColor = 0xFFDCC8;

        private const string NmRowMin = "FUWOA_HL_RowMin";
        private const string NmRowMax = "FUWOA_HL_RowMax";
        private const string NmColMin = "FUWOA_HL_ColMin";
        private const string NmColMax = "FUWOA_HL_ColMax";
        private const string CfPrefix = "FUWOA_HL_";

        private readonly Excel.Application _app;
        private bool _enabled;
        private int _highlightColor;
        private Excel.Worksheet _activeSheet;

        public HighlightManager(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _highlightColor = LoadColorFromRegistry();
        }

        public bool Enabled => _enabled;

        public int HighlightColor
        {
            get => _highlightColor;
            set
            {
                _highlightColor = value;
                SaveColorToRegistry(value);
                if (_enabled) UpdateRuleColors();
            }
        }

        // ── Enable / Disable ──

        public void Enable()
        {
            if (_enabled) return;
            _enabled = true;

            CreateNamedRanges();
            _app.SheetSelectionChange += OnSelectionChange;
            _app.SheetActivate += OnSheetActivate;
            _app.WorkbookActivate += OnWorkbookActivate;

            try
            {
                var sheet = _app.ActiveSheet as Excel.Worksheet;
                if (sheet != null)
                {
                    ApplyFormats(sheet);
                    UpdateNamedRangesFromSelection(sheet, _app.Selection as Excel.Range);
                }
            }
            catch { }
        }

        public void Disable()
        {
            if (!_enabled) return;
            _enabled = false;

            _app.SheetSelectionChange -= OnSelectionChange;
            _app.SheetActivate -= OnSheetActivate;
            _app.WorkbookActivate -= OnWorkbookActivate;

            RemoveAllFormats();
            RemoveNamedRanges();
            _activeSheet = null;
        }

        // ── Events ──

        private void OnSelectionChange(object sh, Excel.Range target)
        {
            if (!_enabled) return;
            var sheet = target?.Worksheet;
            if (sheet != null) UpdateNamedRangesFromSelection(sheet, target);
        }

        private void OnSheetActivate(object sh)
        {
            if (!_enabled) return;
            var sheet = sh as Excel.Worksheet;
            if (sheet == null || sheet == _activeSheet) return;

            RemoveFormatsFromSheet(_activeSheet);
            ApplyFormats(sheet);
            UpdateNamedRangesFromSelection(sheet, _app.Selection as Excel.Range);
        }

        private void OnWorkbookActivate(Excel.Workbook wb)
        {
            if (!_enabled) return;
            var sheet = _app.ActiveSheet as Excel.Worksheet;
            if (sheet == null || sheet == _activeSheet) return;

            RemoveFormatsFromSheet(_activeSheet);
            ApplyFormats(sheet);
            UpdateNamedRangesFromSelection(sheet, _app.Selection as Excel.Range);
        }

        // ── Named Range Helpers ──

        private void EnsureName(string name, int defaultValue)
        {
            try
            {
                var existing = _app.Names.Item(name);
                existing.RefersTo = "=" + defaultValue;
            }
            catch
            {
                _app.Names.Add(name, "=" + defaultValue);
            }
        }

        private void SetNameValue(string name, int value)
        {
            try { _app.Names.Item(name).RefersTo = "=" + value; }
            catch { }
        }

        private void CreateNamedRanges()
        {
            EnsureName(NmRowMin, 0);
            EnsureName(NmRowMax, 0);
            EnsureName(NmColMin, 0);
            EnsureName(NmColMax, 0);
        }

        private void RemoveNamedRanges()
        {
            try
            {
                var names = _app.Names;
                for (int i = names.Count; i >= 1; i--)
                {
                    try
                    {
                        var n = names.Item(i);
                        var nm = n.Name;
                        // 名称可能带前缀如 "Sheet1!FUWOA_HL_RowMin"
                        if (nm == NmRowMin || nm == NmRowMax || nm == NmColMin || nm == NmColMax ||
                            nm.EndsWith("!" + NmRowMin) || nm.EndsWith("!" + NmRowMax) ||
                            nm.EndsWith("!" + NmColMin) || nm.EndsWith("!" + NmColMax))
                            n.Delete();
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void UpdateNamedRangesFromSelection(Excel.Worksheet sheet, Excel.Range selection)
        {
            if (selection == null || sheet == null) return;

            try
            {
                // 合并单元格：MergeArea 对普通单元格返回自身，对合并区域内的单元格返回完整合并范围
                Excel.Range sel;
                try { sel = selection.MergeArea; } catch { sel = selection; }

                int rowMin = int.MaxValue, rowMax = int.MinValue;
                int colMin = int.MaxValue, colMax = int.MinValue;

                foreach (Excel.Range area in sel.Areas)
                {
                    Excel.Range eff;
                    try { eff = area.MergeArea; } catch { eff = area; }
                    int r1 = eff.Row;
                    int r2 = r1 + eff.Rows.Count - 1;
                    int c1 = eff.Column;
                    int c2 = c1 + eff.Columns.Count - 1;

                    if (r1 < rowMin) rowMin = r1;
                    if (r2 > rowMax) rowMax = r2;
                    if (c1 < colMin) colMin = c1;
                    if (c2 > colMax) colMax = c2;
                }

                if (rowMin == int.MaxValue) return;

                // 选中整行或整列时不高亮（含多行/多列：所有列被选中即为整行，所有行被选中即为整列）
                int maxCol = sheet.Columns.Count;
                int maxRow = sheet.Rows.Count;
                if (sel.Columns.Count >= maxCol || sel.Rows.Count >= maxRow)
                    goto ZeroOut;

                // 冻结屏幕刷新，避免命名范围逐个更新时的局部重绘闪烁
                var prevScreenUpdating = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
                try
                {
                    SetNameValue(NmRowMin, rowMin);
                    SetNameValue(NmRowMax, rowMax);
                    SetNameValue(NmColMin, colMin);
                    SetNameValue(NmColMax, colMax);
                }
                finally
                {
                    _app.ScreenUpdating = prevScreenUpdating;
                }

                _activeSheet = sheet;
                return;

            ZeroOut:
                prevScreenUpdating = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
                try
                {
                    SetNameValue(NmRowMin, 0);
                    SetNameValue(NmRowMax, 0);
                    SetNameValue(NmColMin, 0);
                    SetNameValue(NmColMax, 0);
                }
                finally
                {
                    _app.ScreenUpdating = prevScreenUpdating;
                }
                _activeSheet = null;
            }
            catch { }
        }

        // ── Conditional Format Helpers ──

        private static readonly string[] CfFormulas =
        {
            // 0: 行左 — 选中行左侧
            "AND(ROW()>=" + NmRowMin + ", ROW()<=" + NmRowMax
                + ", COLUMN()<" + NmColMin + ")",
            // 1: 行右 — 选中行右侧
            "AND(ROW()>=" + NmRowMin + ", ROW()<=" + NmRowMax
                + ", COLUMN()>" + NmColMax + ")",
            // 2: 列上 — 选中列上方
            "AND(COLUMN()>=" + NmColMin + ", COLUMN()<=" + NmColMax
                + ", ROW()<" + NmRowMin + ")",
            // 3: 列下 — 选中列下方
            "AND(COLUMN()>=" + NmColMin + ", COLUMN()<=" + NmColMax
                + ", ROW()>" + NmRowMax + ")",
        };

        private void ApplyFormats(Excel.Worksheet sheet)
        {
            try
            {
                // 全表应用条件格式，避免选中数据区外单元格时被 UsedRange 裁剪
                var rng = sheet.Cells;
                var fcs = rng.FormatConditions;
                int color = _highlightColor;
                foreach (var formula in CfFormulas)
                {
                    var fc = fcs.Add(
                        Excel.XlFormatConditionType.xlExpression,
                        Type.Missing,
                        "=" + formula);
                    fc.Interior.Color = color;
                    fc.StopIfTrue = false;
                }
                _activeSheet = sheet;
            }
            catch { }
        }

        private void RemoveFormatsFromSheet(Excel.Worksheet sheet)
        {
            if (sheet == null) return;
            try
            {
                var fcs = sheet.Cells.FormatConditions;
                for (int i = fcs.Count; i >= 1; i--)
                {
                    try
                    {
                        var fc = fcs[i];
                        var f1 = fc.Formula1 as string;
                        if (f1 != null && f1.Contains(CfPrefix))
                            fc.Delete();
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void RemoveAllFormats()
        {
            try
            {
                foreach (Excel.Worksheet sheet in _app.Worksheets)
                    RemoveFormatsFromSheet(sheet);
            }
            catch { }
            RemoveFormatsFromSheet(_activeSheet);
        }

        private void UpdateRuleColors()
        {
            if (_activeSheet == null) return;
            try
            {
                var fcs = _activeSheet.Cells.FormatConditions;
                for (int i = 1; i <= fcs.Count; i++)
                {
                    try
                    {
                        var fc = fcs[i];
                        var f1 = fc.Formula1 as string;
                        if (f1 != null && f1.Contains(CfPrefix))
                            fc.Interior.Color = _highlightColor;
                    }
                    catch { }
                }
            }
            catch { }
        }

        // ── Registry ──

        private static int LoadColorFromRegistry()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegPath);
                if (key != null)
                {
                    var val = key.GetValue(ColorKey);
                    if (val is int i) return i;
                }
            }
            catch { }
            return DefaultColor;
        }

        private static void SaveColorToRegistry(int color)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RegPath);
                key.SetValue(ColorKey, color, RegistryValueKind.DWord);
            }
            catch { }
        }
    }
}
#endif
