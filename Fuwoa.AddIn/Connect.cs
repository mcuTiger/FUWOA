using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

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

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            _applicationObject = application;
            _addInInstance = addInInst;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _ribbonUI = null;
            _applicationObject = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

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
            xml.AppendLine($"                  label=\"{E(export)}\"");
            xml.AppendLine($"                  screentip=\"{E(screentip)}\"");
            xml.AppendLine($"                  supertip=\"{E(supertip)}\"");
            xml.AppendLine( "                  onAction=\"OnExportCountClick\"");
            xml.AppendLine( "                  imageMso=\"CreateReportFromWizard\"");
            xml.AppendLine( "                  size=\"large\"/>");
            xml.AppendLine( "        </group>");

            // About group
            xml.AppendLine($"        <group id=\"AboutGroup\" label=\"{E(about)}\">");
            xml.AppendLine($"          <labelControl id=\"VersionLabel\" label=\"{E(version)}\"/>");

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
            command.Execute(_applicationObject as Excel.Application);
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
    }
}
