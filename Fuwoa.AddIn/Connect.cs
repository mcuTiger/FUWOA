using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Fuwoa.Core.ExportCount;

namespace Fuwoa.AddIn
{
    [ComVisible(true)]
    [Guid("E7A8B9C0-D1E2-4F3A-8B6C-9D0E1F2A3B4C")]
    [ProgId("Fuwoa.AddIn")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _applicationObject;
        private object _addInInstance;
        private IRibbonUI _ribbonUI;
        private Timer _filterTimer;
        private bool _lastFilterMode;
        private Excel.AppEvents_SheetActivateEventHandler _sheetActivateHandler;
        private static SortMode _sortMode;
        private static bool _sortModeLoaded;
        private static bool _sortDescending = true;
        private static bool _sortDescLoaded;
        private static bool _showPercentage;
        private static bool _showPerLoaded;
        private HighlightManager _highlightManager;

        private static SortMode SortMode
        {
            get
            {
                if (!_sortModeLoaded)
                {
                    _sortModeLoaded = true;
                    try
                    {
                        using var key = Microsoft.Win32.Registry.CurrentUser
                            .OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                        _sortMode = (key?.GetValue("SortMode") as string) == "Title"
                            ? SortMode.ByTitle : SortMode.ByCount;
                    }
                    catch { _sortMode = SortMode.ByCount; }
                }
                return _sortMode;
            }
            set
            {
                _sortMode = value;
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser
                        .CreateSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                    key.SetValue("SortMode", value == SortMode.ByTitle ? "Title" : "Count");
                }
                catch { }
            }
        }

        private static bool SortDescending
        {
            get
            {
                if (!_sortDescLoaded)
                {
                    _sortDescLoaded = true;
                    try
                    {
                        using var key = Microsoft.Win32.Registry.CurrentUser
                            .OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                        _sortDescending = (key?.GetValue("SortOrder") as string) != "Asc";
                    }
                    catch { _sortDescending = true; }
                }
                return _sortDescending;
            }
            set
            {
                _sortDescending = value;
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser
                        .CreateSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                    key.SetValue("SortOrder", value ? "Desc" : "Asc");
                }
                catch { }
            }
        }

        private static bool ShowPercentage
        {
            get
            {
                if (!_showPerLoaded)
                {
                    _showPerLoaded = true;
                    try
                    {
                        using var key = Microsoft.Win32.Registry.CurrentUser
                            .OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                        _showPercentage = (key?.GetValue("ShowPercentage") as int? ?? 0) != 0;
                    }
                    catch { _showPercentage = false; }
                }
                return _showPercentage;
            }
            set
            {
                _showPercentage = value;
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser
                        .CreateSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn");
                    key.SetValue("ShowPercentage", value ? 1 : 0);
                }
                catch { }
            }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            _applicationObject = application;
            _addInInstance = addInInst;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            StopFilterWatcher();
            _highlightManager?.Disable();
            _ribbonUI = null;
            _applicationObject = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            StartFilterWatcher();
            _highlightManager = new HighlightManager(_applicationObject as Excel.Application);
        }

        public void OnBeginShutdown(ref Array custom)
        {
            StopFilterWatcher();
        }

        public string GetCustomUI(string ribbonID)
        {
            string export     = L("exportCount");
            string screentip  = L("exportCountScreentip");
            string supertip   = L("exportCountSupertip");
            string dataTools  = L("dataTools");
            string about      = L("about");
            string version    = L("version");
            string language   = L("language");

            var xml = new StringBuilder();
            xml.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            xml.AppendLine("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">");
            xml.AppendLine("  <ribbon>");
            xml.AppendLine("    <tabs>");
            xml.AppendLine("      <tab id=\"FuwoaTab\" label=\"FUWOA\" insertAfterMso=\"TabHome\">");

            // Data Tools group
            xml.AppendLine($"        <group id=\"DataToolsGroup\" label=\"{E(dataTools)}\">");
            xml.AppendLine($"          <button id=\"ExportCountBtn\"");
            xml.AppendLine( "                  getLabel=\"GetExportCountLabel\"");
            xml.AppendLine($"                  screentip=\"{E(screentip)}\"");
            xml.AppendLine($"                  supertip=\"{E(supertip)}\"");
            xml.AppendLine( "                  onAction=\"OnExportCountClick\"");
            xml.AppendLine( "                  imageMso=\"CreateReportFromWizard\"");
            xml.AppendLine( "                  size=\"large\"/>");
            xml.AppendLine($"          <box id=\"SortBox\" boxStyle=\"vertical\">");
            xml.AppendLine($"          <dropDown id=\"SortDropDown\"");
            xml.AppendLine($"                    label=\"{E(L("sortBy"))}\"");
            xml.AppendLine( "                    sizeString=\"WWWWWWW\"");
            xml.AppendLine( "                    getSelectedItemIndex=\"GetSortSelectedIndex\"");
            xml.AppendLine( "                    onAction=\"OnSortDropDownAction\">");
            xml.AppendLine($"            <item id=\"SortByCount\" label=\"{E(L("sortByCount"))}\"/>");
            xml.AppendLine($"            <item id=\"SortByTitle\" label=\"{E(L("sortByTitle"))}\"/>");
            xml.AppendLine( "          </dropDown>");
            xml.AppendLine($"          <dropDown id=\"OrderDropDown\"");
            xml.AppendLine($"                    label=\"{E(L("sortOrder"))}\"");
            xml.AppendLine( "                    sizeString=\"WWWWWWW\"");
            xml.AppendLine( "                    getSelectedItemIndex=\"GetOrderSelectedIndex\"");
            xml.AppendLine( "                    onAction=\"OnOrderDropDownAction\">");
            xml.AppendLine($"            <item id=\"OrderDesc\" label=\"{E(L("sortDesc"))}\"/>");
            xml.AppendLine($"            <item id=\"OrderAsc\"  label=\"{E(L("sortAsc"))}\"/>");
            xml.AppendLine( "          </dropDown>");
            xml.AppendLine($"          <toggleButton id=\"PercentageToggle\"");
            xml.AppendLine( "                       getLabel=\"GetPercentageToggleLabel\"");
            xml.AppendLine( "                       getPressed=\"GetPercentageTogglePressed\"");
            xml.AppendLine( "                       onAction=\"OnPercentageToggleAction\"");
            xml.AppendLine( "                       imageMso=\"PercentStyle\"/>");
            xml.AppendLine( "          </box>");
            xml.AppendLine($"          <separator id=\"ExportSep\"/>");
            xml.AppendLine($"          <button id=\"SplitByColumnBtn\"");
            xml.AppendLine( "                  getLabel=\"GetSplitByColumnLabel\"");
            xml.AppendLine($"                  screentip=\"{E(L("splitByColumnScreentip"))}\"");
            xml.AppendLine($"                  supertip=\"{E(L("splitByColumnSupertip"))}\"");
            xml.AppendLine( "                  onAction=\"OnSplitByColumnClick\"");
            xml.AppendLine( "                  imageMso=\"GroupField\"");
            xml.AppendLine( "                  size=\"large\"/>");
            xml.AppendLine($"          <menu id=\"MarkdownImportMenu\"");
            xml.AppendLine( "                  getLabel=\"GetMarkdownImportLabel\"");
            xml.AppendLine($"                  screentip=\"{E(L("mdImportScreentip"))}\"");
            xml.AppendLine($"                  supertip=\"{E(L("mdImportSupertip"))}\"");
            xml.AppendLine( "                  imageMso=\"PasteAsNestedTable\"");
            xml.AppendLine( "                  size=\"large\">");
            xml.AppendLine($"            <button id=\"MdImportFileBtn\"");
            xml.AppendLine( "                    label=\"" + E(L("mdImportFile")) + "\"");
            xml.AppendLine( "                    onAction=\"OnMdImportFileClick\"");
            xml.AppendLine( "                    imageMso=\"OpenFile\"/>");
            xml.AppendLine($"            <button id=\"MdImportClipboardBtn\"");
            xml.AppendLine( "                    label=\"" + E(L("mdImportClipboard")) + "\"");
            xml.AppendLine( "                    onAction=\"OnMdImportClipboardClick\"");
            xml.AppendLine( "                    imageMso=\"Copy\"/>");
            xml.AppendLine( "          </menu>");
            xml.AppendLine( "        </group>");

            // Visual Tools group
            xml.AppendLine($"        <group id=\"HighlightGroup\" label=\"{E(L("visualTools"))}\">");
            xml.AppendLine($"          <toggleButton id=\"HighlightToggle\"");
            xml.AppendLine( "                       getLabel=\"GetHighlightToggleLabel\"");
            xml.AppendLine( "                       getPressed=\"GetHighlightTogglePressed\"");
            xml.AppendLine( "                       onAction=\"OnHighlightToggleAction\"");
            xml.AppendLine( "                       imageMso=\"TextHighlightColorPicker\"");
            xml.AppendLine( "                       size=\"large\"/>");
            xml.AppendLine($"          <gallery id=\"HighlightColorGallery\"");
            xml.AppendLine( "                       getLabel=\"GetHighlightColorLabel\"");
            xml.AppendLine( "                       getItemCount=\"GetHighlightColorItemCount\"");
            xml.AppendLine( "                       getItemImage=\"GetHighlightColorItemImage\"");
            xml.AppendLine( "                       onAction=\"OnHighlightColorAction\"");
            xml.AppendLine( "                       columns=\"5\"");
            xml.AppendLine( "                       itemWidth=\"20\"");
            xml.AppendLine( "                       itemHeight=\"20\"/>");
            xml.AppendLine( "        </group>");

            // About group (always last)
            xml.AppendLine($"        <group id=\"AboutGroup\" label=\"{E(about)}\">");
            xml.AppendLine($"          <labelControl id=\"VersionLabel\" label=\"{E(version)}\"/>");
            // Development / BETA mode tag (visible in debug builds, hidden in Release/MSI builds)
#if !RELEASE
            var devTag = L("devTag");
            if (!string.IsNullOrEmpty(devTag))
                xml.AppendLine($"          <labelControl id=\"DevTagLabel\" label=\"{E(devTag)}\"/>");
#endif

            // Separator row: icon + label before dropdown
            xml.AppendLine( "          <box id=\"LangBox\" boxStyle=\"horizontal\">");
            xml.AppendLine( "            <button id=\"LangIcon\"");
            xml.AppendLine( "                    imageMso=\"Translate\"");
            xml.AppendLine( "                    enabled=\"false\"");
            xml.AppendLine( "                    showLabel=\"false\"");
            xml.AppendLine( "                    size=\"normal\"/>");
            xml.AppendLine($"            <dropDown id=\"LangDropDown\"");
            xml.AppendLine($"                      label=\"{E(language)}\"");
            xml.AppendLine( "                      getSelectedItemIndex=\"GetLangSelectedIndex\"");
            xml.AppendLine( "                      onAction=\"OnLanguageDropDownAction\">");
            xml.AppendLine( "              <item id=\"LangZhCN\" label=\"简体中文\"/>");
            xml.AppendLine( "              <item id=\"LangZhTW\" label=\"繁體中文\"/>");
            xml.AppendLine( "              <item id=\"LangEn\"   label=\"English\"/>");
            xml.AppendLine( "              <item id=\"LangDe\"   label=\"Deutsch\"/>");
            xml.AppendLine( "              <item id=\"LangFr\"   label=\"Français\"/>");
            xml.AppendLine( "              <item id=\"LangRu\"   label=\"Русский\"/>");
            xml.AppendLine( "              <item id=\"LangVi\"   label=\"Tiếng Việt\"/>");
            xml.AppendLine( "              <item id=\"LangTh\"   label=\"ไทย\"/>");
            xml.AppendLine( "              <item id=\"LangJa\"   label=\"日本語\"/>");
            xml.AppendLine( "              <item id=\"LangId\"   label=\"Bahasa Indonesia\"/>");
            xml.AppendLine( "              <item id=\"LangBo\"   label=\"བོད་སྐད།\"/>");
            xml.AppendLine( "              <item id=\"LangUg\"   label=\"ئۇيغۇرچە\"/>");
            xml.AppendLine( "            </dropDown>");
            xml.AppendLine( "          </box>");

            xml.AppendLine( "        </group>");

            xml.AppendLine("      </tab>");
            xml.AppendLine("    </tabs>");
            xml.AppendLine("  </ribbon>");
            xml.AppendLine("</customUI>");

            return xml.ToString();
        }

        public void OnLoad(IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
        }

        // ── Export Count ──

        public void OnExportCountClick(IRibbonControl control)
        {
            var command = new Commands.ExportCountCommand();
            command.Execute(_applicationObject as Excel.Application, SortMode, SortDescending, ShowPercentage);
        }

        public void OnSplitByColumnClick(IRibbonControl control)
        {
            var command = new Commands.SplitByColumnCommand();
            command.Execute(_applicationObject as Excel.Application);
        }

        // ── Markdown Import ──

        public string GetMarkdownImportLabel(IRibbonControl control)
        {
            try { return L("mdImport"); }
            catch { return "Markdown Import"; }
        }

        public void OnMdImportFileClick(IRibbonControl control)
        {
            var command = new Commands.MarkdownImportCommand();
            command.Execute(_applicationObject as Excel.Application, Commands.MarkdownImportSource.File);
        }

        public void OnMdImportClipboardClick(IRibbonControl control)
        {
            var command = new Commands.MarkdownImportCommand();
            command.Execute(_applicationObject as Excel.Application, Commands.MarkdownImportSource.Clipboard);
        }

        public string GetSplitByColumnLabel(IRibbonControl control)
        {
            try
            {
                var app = _applicationObject as Excel.Application;
                var sheet = app?.ActiveSheet as Excel.Worksheet;
                if (sheet != null && sheet.AutoFilterMode)
                {
                    var filter = sheet.AutoFilter;
                    if (filter != null && filter.FilterMode)
                        return L("splitByColumnFiltered");
                }
            }
            catch { }
            return L("splitByColumnAll");
        }

        public string GetExportCountLabel(IRibbonControl control)
        {
            try
            {
                var app = _applicationObject as Excel.Application;
                var sheet = app?.ActiveSheet as Excel.Worksheet;
                if (sheet != null && sheet.AutoFilterMode)
                {
                    var filter = sheet.AutoFilter;
                    if (filter != null && filter.FilterMode)
                        return L("exportCountFiltered");
                }
            }
            catch { }
            return L("exportCountAll");
        }

        // ── Sort Dropdown ──

        public int GetSortSelectedIndex(IRibbonControl c)
        {
            return SortMode == SortMode.ByTitle ? 1 : 0;
        }

        public void OnSortDropDownAction(IRibbonControl control,
            string selectedId, int selectedIndex)
        {
            SortMode = selectedId == "SortByTitle"
                ? SortMode.ByTitle : SortMode.ByCount;
        }

        // ── Order Dropdown ──

        public int GetOrderSelectedIndex(IRibbonControl c)
        {
            return SortDescending ? 0 : 1;
        }

        public void OnOrderDropDownAction(IRibbonControl control,
            string selectedId, int selectedIndex)
        {
            SortDescending = selectedId != "OrderAsc";
        }

        // ── Percentage Toggle ──

        public bool GetPercentageTogglePressed(IRibbonControl control)
        {
            return ShowPercentage;
        }

        public void OnPercentageToggleAction(IRibbonControl control, bool pressed)
        {
            ShowPercentage = pressed;
        }

        public string GetPercentageToggleLabel(IRibbonControl control)
        {
            try { return L("showPercentage"); }
            catch { return "%"; }
        }

        public string GetDropDownSizeString(IRibbonControl control)
        {
            var lang = LanguageManager.Current;
            if (lang == Language.zh_CN || lang == Language.zh_TW) return "WWWWWW";
            if (lang == Language.ja) return "WWWWWWW";
            if (lang == Language.de || lang == Language.ru) return "WWWWWWWW";
            if (lang == Language.th) return "WWWWWWWWW";
            return "WWWWWWW";
        }

        // ── Language Dropdown ──

        public int GetLangSelectedIndex(IRibbonControl c)
        {
            try
            {
                return LanguageManager.Current switch
                {
                    Language.zh_TW => 1,
                    Language.en    => 2,
                    Language.de    => 3,
                    Language.fr    => 4,
                    Language.ru    => 5,
                    Language.vi    => 6,
                    Language.th    => 7,
                    Language.ja    => 8,
                    Language.id    => 9,
                    Language.bo    => 10,
                    Language.ug    => 11,
                    _ => 0
                };
            }
            catch { return 0; }
        }

        public void OnLanguageDropDownAction(IRibbonControl control,
            string selectedId, int selectedIndex)
        {
            try
            {
                Language lang = selectedId switch
                {
                    "LangZhTW" => Language.zh_TW,
                    "LangEn"   => Language.en,
                    "LangDe"   => Language.de,
                    "LangFr"   => Language.fr,
                    "LangRu"   => Language.ru,
                    "LangVi"   => Language.vi,
                    "LangTh"   => Language.th,
                    "LangJa"   => Language.ja,
                    "LangId"   => Language.id,
                    "LangBo"   => Language.bo,
                    "LangUg"   => Language.ug,
                    _ => Language.zh_CN
                };

                if (lang == LanguageManager.Current) return;

                LanguageManager.SetLanguage(lang);
                _ribbonUI?.Invalidate();

                string msg = L("langRestart");
                if (string.IsNullOrEmpty(msg) || msg == "langRestart")
                    msg = "Language changed. Please restart Excel for the change to take effect.";

                MessageBox.Show(msg, "FUWOA", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch { }
        }

        // ── Highlight Toggle ──

        public string GetHighlightToggleLabel(IRibbonControl control)
        {
            try { return L("highlightToggle"); } catch { return "Highlight"; }
        }

        public bool GetHighlightTogglePressed(IRibbonControl control)
        {
            return _highlightManager?.Enabled ?? false;
        }

        public void OnHighlightToggleAction(IRibbonControl control, bool pressed)
        {
            if (pressed)
                _highlightManager?.Enable();
            else
                _highlightManager?.Disable();
        }

        // ── Highlight Color Gallery ──

        private static readonly Dictionary<int, object> ColorImageCache =
            new Dictionary<int, object>();

        private static readonly int[] PresetColors =
        {
            0xE0FFFF, // Light Yellow
            0xFFDCC8, // Light Blue (default)
            0xE0FFE0, // Light Green
            0xFFE0E0, // Light Pink
            0xDDDDFF, // Light Purple
            0xFFFFCC, // Pale Yellow
            0xE8E8E8, // Light Gray
            0x0000FF, // Red
            0x0080FF, // Orange
            0x00FFFF, // Yellow
            0x00FF00, // Lime
            0xFF0000, // Blue
            0x800080, // Purple
            0x800000, // Navy
        };

        public int GetHighlightColorItemCount(IRibbonControl control) => PresetColors.Length;

        public string GetHighlightColorLabel(IRibbonControl control)
        {
            try { return L("highlightColor"); } catch { return "Color"; }
        }

        public object GetHighlightColorItemImage(IRibbonControl control, int index)
        {
            if (index < 0 || index >= PresetColors.Length) return null;
            int ole = PresetColors[index];
            if (ColorImageCache.TryGetValue(ole, out var cached)) return cached;
            var color = ColorTranslator.FromOle(ole);
            var bmp = new Bitmap(16, 16);
            try
            {
                using var g = Graphics.FromImage(bmp);
                using var brush = new SolidBrush(color);
                g.FillRectangle(brush, 0, 0, 16, 16);
                using var pen = new Pen(Color.DarkGray);
                g.DrawRectangle(pen, 0, 0, 15, 15);
                var pic = PictureDispConverter.FromBitmap(bmp);
                ColorImageCache[ole] = pic;
                return pic;
            }
            catch { bmp.Dispose(); return null; }
        }

        public void OnHighlightColorAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < PresetColors.Length && _highlightManager != null)
                _highlightManager.HighlightColor = PresetColors[selectedIndex];
        }

        /// <summary>Bitmap → IPictureDisp 转换器（通过 AxHost 受保护方法）</summary>
        private class PictureDispConverter : System.Windows.Forms.AxHost
        {
            private PictureDispConverter() : base("{00000000-0000-0000-0000-000000000000}") { }
            public static object FromBitmap(Bitmap bmp) => GetIPictureDispFromPicture(bmp);
        }

        // ── Helpers ──

        private static string L(string key)
        {
            try { return LanguageManager.Get(key); }
            catch { return key; }
        }

        private static string E(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;")
                    .Replace("\"", "&quot;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("'", "&apos;");
        }

        // ── Filter mode watcher for real-time label update ──

        private void StartFilterWatcher()
        {
            try
            {
                _filterTimer = new Timer { Interval = 500 };
                _filterTimer.Tick += (s, e) =>
                {
                    try
                    {
                        var app = _applicationObject as Excel.Application;
                        var sheet = app?.ActiveSheet as Excel.Worksheet;
                        bool current = sheet != null && sheet.AutoFilterMode &&
                                       sheet.AutoFilter != null &&
                                       sheet.AutoFilter.FilterMode;
                        if (current != _lastFilterMode)
                        {
                            _lastFilterMode = current;
                            _ribbonUI?.InvalidateControl("ExportCountBtn");
                            _ribbonUI?.InvalidateControl("SplitByColumnBtn");
                        }
                    }
                    catch { }
                };
                _filterTimer.Start();

                // Also hook SheetActivate for sheet switches
                var app2 = _applicationObject as Excel.Application;
                if (app2 != null)
                {
                    _sheetActivateHandler = sh =>
                    {
                        _ribbonUI?.InvalidateControl("ExportCountBtn");
                        _ribbonUI?.InvalidateControl("SplitByColumnBtn");
                    };
                    app2.SheetActivate += _sheetActivateHandler;
                }
            }
            catch { }
        }

        private void StopFilterWatcher()
        {
            try
            {
                _filterTimer?.Stop();
                _filterTimer?.Dispose();
                _filterTimer = null;
            }
            catch { }
            if (_applicationObject is Excel.Application app && _sheetActivateHandler != null)
            {
                try { app.SheetActivate -= _sheetActivateHandler; } catch { }
                _sheetActivateHandler = null;
            }
        }
    }
}
