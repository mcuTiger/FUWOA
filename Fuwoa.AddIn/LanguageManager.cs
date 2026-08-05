using Microsoft.Win32;
using System.Collections.Generic;

namespace Fuwoa.AddIn
{
    public enum Language
    {
        zh_CN, zh_TW, en,
        de, fr, ru,
        vi, th, id,
        bo, ug, ja
    }

    public static class LanguageManager
    {
        private const string RegPath = @"SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn";
        private const string LangKey = "Language";

        private static Language _current;
        private static bool _loaded;

        private static readonly Dictionary<Language, Dictionary<string, string>> Strings =
            new Dictionary<Language, Dictionary<string, string>>
        {
            [Language.zh_CN] = new Dictionary<string, string>
            {
                ["exportCount"] = "导出计数",
                ["exportCountScreentip"] = "导出唯一值及出现次数",
                ["exportCountSupertip"] = "选中某一列的标题单元格，点击将下方该列所有唯一值及其出现次数按降序导出到新工作表。",
                ["dataTools"] = "数据工具",
                ["about"] = "关于",
                ["version"] = "FUWOA BETA",
                ["language"] = "语言",
                ["langRestart"] = "部分翻译将在重启 Excel 后生效。",
                ["count"] = "计数",
                ["column"] = "列",
                ["noExcelApp"] = "无法获取 Excel 应用程序实例。",
                ["selectOneCell"] = "请选中某一列的标题单元格（单个单元格）。",
                ["noDataBelow"] = "该列标题下方没有数据。",
                ["exportFailed"] = "导出计数失败",
                ["devTag"] = "当前是 Tiger 的开发测试版本",
                ["exportCountAll"] = "导出计数 (全部)",
                ["exportCountFiltered"] = "导出计数 (已筛选)",
                ["sortBy"] = "排序",
                ["sortByCount"] = "按数值",
                ["sortByTitle"] = "按标题",
                ["sortOrder"] = "方向",
                ["sortDesc"] = "降序",
                ["sortAsc"] = "升序",
                ["highlightToggle"] = "行列高亮",
                ["highlightColor"] = "高亮颜色",
            },
            [Language.zh_TW] = new Dictionary<string, string>
            {
                ["exportCount"] = "匯出計數",
                ["exportCountScreentip"] = "匯出唯一值及出現次數",
                ["exportCountSupertip"] = "選取某一欄的標題儲存格，點擊將下方該欄所有唯一值及其出現次數按降序匯出到新工作表。",
                ["dataTools"] = "資料工具",
                ["about"] = "關於",
                ["version"] = "FUWOA BETA",
                ["language"] = "語言",
                ["langRestart"] = "部分翻譯將在重啟 Excel 後生效。",
                ["count"] = "計數",
                ["column"] = "欄",
                ["noExcelApp"] = "無法取得 Excel 應用程式執行個體。",
                ["selectOneCell"] = "請選取某一欄的標題儲存格（單一儲存格）。",
                ["noDataBelow"] = "該欄標題下方沒有資料。",
                ["exportFailed"] = "匯出計數失敗",
                ["devTag"] = "",
                ["exportCountAll"] = "匯出計數 (全部)",
                ["exportCountFiltered"] = "匯出計數 (已篩選)",
                ["sortBy"] = "排序",
                ["sortByCount"] = "按數值",
                ["sortByTitle"] = "按標題",
                ["sortOrder"] = "方向",
                ["sortDesc"] = "降序",
                ["sortAsc"] = "升序",
                ["visualTools"] = "視覺工具",
                ["highlightToggle"] = "行列醒目提示",
            },
            [Language.en] = new Dictionary<string, string>
            {
                ["exportCount"] = "Export Count",
                ["exportCountScreentip"] = "Export unique values with counts",
                ["exportCountSupertip"] = "Select a header cell in any column, then click to export all unique values and their occurrence counts to a new worksheet, sorted by count descending.",
                ["dataTools"] = "Data Tools",
                ["about"] = "About",
                ["version"] = "FUWOA BETA",
                ["language"] = "Language",
                ["langRestart"] = "Some translations require Excel restart.",
                ["count"] = "Count",
                ["column"] = "Col",
                ["noExcelApp"] = "Unable to get Excel application instance.",
                ["selectOneCell"] = "Please select a single header cell in any column.",
                ["noDataBelow"] = "No data found below the header cell.",
                ["exportFailed"] = "Export Count Failed",
                ["devTag"] = "Tiger's dev/testing build",
                ["exportCountAll"] = "Export Count (All)",
                ["exportCountFiltered"] = "Export Count (Filtered)",
                ["sortBy"] = "Sort",
                ["sortByCount"] = "By Count",
                ["sortByTitle"] = "By Title",
                ["sortOrder"] = "Order",
                ["sortDesc"] = "Descending",
                ["sortAsc"] = "Ascending",
                ["visualTools"] = "Visual Tools",
                ["highlightToggle"] = "Row/Col Highlight",
                ["highlightColor"] = "Highlight Color",
            },
            [Language.de] = new Dictionary<string, string>
            {
                ["exportCount"] = "Zählung exportieren",
                ["exportCountScreentip"] = "Eindeutige Werte mit Häufigkeit exportieren",
                ["exportCountSupertip"] = "Wählen Sie eine Kopfzeile in einer beliebigen Spalte und klicken Sie dann, um alle eindeutigen Werte und deren Häufigkeit absteigend in ein neues Arbeitsblatt zu exportieren.",
                ["dataTools"] = "Datentools",
                ["about"] = "Über",
                ["version"] = "FUWOA BETA",
                ["language"] = "Sprache",
                ["langRestart"] = "Einige Übersetzungen erfordern einen Excel-Neustart.",
                ["count"] = "Anzahl",
                ["column"] = "Spalte",
                ["noExcelApp"] = "Excel-Anwendungsinstanz kann nicht abgerufen werden.",
                ["selectOneCell"] = "Bitte wählen Sie eine einzelne Kopfzelle in einer Spalte.",
                ["noDataBelow"] = "Keine Daten unterhalb der Kopfzelle gefunden.",
                ["exportFailed"] = "Export fehlgeschlagen",
                ["devTag"] = "",
                ["exportCountAll"] = "Zählung exportieren (Alle)",
                ["exportCountFiltered"] = "Zählung exportieren (Gefiltert)",
                ["sortBy"] = "Sortierung",
                ["sortByCount"] = "Nach Anzahl",
                ["sortByTitle"] = "Nach Titel",
                ["sortOrder"] = "Reihenfolge",
                ["sortDesc"] = "Absteigend",
                ["sortAsc"] = "Aufsteigend",
                ["visualTools"] = "Visuelle Werkzeuge",
                ["highlightToggle"] = "Zeilen/Spalten hervorheben",
                ["highlightColor"] = "Hervorhebungsfarbe"
            },
            [Language.fr] = new Dictionary<string, string>
            {
                ["exportCount"] = "Exporter le décompte",
                ["exportCountScreentip"] = "Exporter les valeurs uniques avec leur occurrence",
                ["exportCountSupertip"] = "Sélectionnez une cellule d'en-tête dans n'importe quelle colonne, puis cliquez pour exporter toutes les valeurs uniques et leur nombre d'occurrences dans une nouvelle feuille, triées par ordre décroissant.",
                ["dataTools"] = "Outils de données",
                ["about"] = "À propos",
                ["version"] = "FUWOA BETA",
                ["language"] = "Langue",
                ["langRestart"] = "Certaines traductions nécessitent un redémarrage d'Excel.",
                ["count"] = "Nombre",
                ["column"] = "Colonne",
                ["noExcelApp"] = "Impossible d'obtenir l'instance de l'application Excel.",
                ["selectOneCell"] = "Veuillez sélectionner une seule cellule d'en-tête.",
                ["noDataBelow"] = "Aucune donnée sous la cellule d'en-tête.",
                ["exportFailed"] = "Échec de l'exportation",
                ["devTag"] = "",
                ["exportCountAll"] = "Exporter le décompte (Tout)",
                ["exportCountFiltered"] = "Exporter le décompte (Filtré)",
                ["sortBy"] = "Tri",
                ["sortByCount"] = "Par nombre",
                ["sortByTitle"] = "Par titre",
                ["sortOrder"] = "Ordre",
                ["sortDesc"] = "Décroissant",
                ["sortAsc"] = "Croissant",
                ["visualTools"] = "Outils visuels",
                ["highlightToggle"] = "Surligner lignes/colonnes",
                ["highlightColor"] = "Couleur de surbrillance"
            },
            [Language.ru] = new Dictionary<string, string>
            {
                ["exportCount"] = "Экспорт подсчёта",
                ["exportCountScreentip"] = "Экспорт уникальных значений с количеством",
                ["exportCountSupertip"] = "Выберите ячейку заголовка в любом столбце и нажмите, чтобы экспортировать все уникальные значения и их количество в новый лист, отсортированные по убыванию.",
                ["dataTools"] = "Инструменты данных",
                ["about"] = "О программе",
                ["version"] = "FUWOA BETA",
                ["language"] = "Язык",
                ["langRestart"] = "Часть перевода вступит в силу после перезапуска Excel.",
                ["count"] = "Количество",
                ["column"] = "Столбец",
                ["noExcelApp"] = "Не удалось получить экземпляр приложения Excel.",
                ["selectOneCell"] = "Выберите одну ячейку заголовка в столбце.",
                ["noDataBelow"] = "Нет данных под ячейкой заголовка.",
                ["exportFailed"] = "Ошибка экспорта",
                ["devTag"] = "",
                ["exportCountAll"] = "Экспорт подсчёта (Всё)",
                ["exportCountFiltered"] = "Экспорт подсчёта (Отфильтровано)",
                ["sortBy"] = "Сортировка",
                ["sortByCount"] = "По количеству",
                ["sortByTitle"] = "По заголовку",
                ["sortOrder"] = "Порядок",
                ["sortDesc"] = "По убыванию",
                ["sortAsc"] = "По возрастанию",
                ["highlightToggle"] = "Подсветка строк и столбцов",
                ["highlightColor"] = "Цвет подсветки"
            },
            [Language.vi] = new Dictionary<string, string>
            {
                ["exportCount"] = "Xuất thống kê",
                ["exportCountScreentip"] = "Xuất giá trị duy nhất kèm số lần xuất hiện",
                ["exportCountSupertip"] = "Chọn một ô tiêu đề trong cột bất kỳ, sau đó nhấn để xuất tất cả giá trị duy nhất và số lần xuất hiện sang sheet mới, sắp xếp giảm dần.",
                ["dataTools"] = "Công cụ dữ liệu",
                ["about"] = "Giới thiệu",
                ["version"] = "FUWOA BETA",
                ["language"] = "Ngôn ngữ",
                ["langRestart"] = "Một số bản dịch cần khởi động lại Excel.",
                ["count"] = "Số lượng",
                ["column"] = "Cột",
                ["noExcelApp"] = "Không thể lấy phiên bản ứng dụng Excel.",
                ["selectOneCell"] = "Vui lòng chọn một ô tiêu đề duy nhất trong cột.",
                ["noDataBelow"] = "Không có dữ liệu bên dưới ô tiêu đề.",
                ["exportFailed"] = "Xuất thất bại",
                ["devTag"] = "",
                ["exportCountAll"] = "Xuất thống kê (Tất cả)",
                ["exportCountFiltered"] = "Xuất thống kê (Đã lọc)",
                ["sortBy"] = "Sắp xếp",
                ["sortByCount"] = "Theo số lượng",
                ["sortByTitle"] = "Theo tiêu đề",
                ["sortOrder"] = "Thứ tự",
                ["sortDesc"] = "Giảm dần",
                ["sortAsc"] = "Tăng dần",
                ["visualTools"] = "Công cụ trực quan",
                ["highlightToggle"] = "Đánh dấu hàng/cột",
                ["highlightColor"] = "Màu đánh dấu"
            },
            [Language.th] = new Dictionary<string, string>
            {
                ["exportCount"] = "ส่งออกจำนวน",
                ["exportCountScreentip"] = "ส่งออกค่าที่ไม่ซ้ำพร้อมจำนวน",
                ["exportCountSupertip"] = "เลือกเซลล์ส่วนหัวในคอลัมน์ใดก็ได้ จากนั้นคลิกเพื่อส่งออกค่าที่ไม่ซ้ำทั้งหมดและจำนวนที่ปรากฏไปยังแผ่นงานใหม่ โดยเรียงตามจำนวนจากมากไปน้อย",
                ["dataTools"] = "เครื่องมือข้อมูล",
                ["about"] = "เกี่ยวกับ",
                ["version"] = "FUWOA BETA",
                ["language"] = "ภาษา",
                ["langRestart"] = "การแปลบางส่วนจะมีผลหลังจากรีสตาร์ท Excel",
                ["count"] = "จำนวน",
                ["column"] = "คอลัมน์",
                ["noExcelApp"] = "ไม่สามารถเข้าถึงอินสแตนซ์ของแอปพลิเคชัน Excel",
                ["selectOneCell"] = "โปรดเลือกเซลล์ส่วนหัวเพียงเซลล์เดียวในคอลัมน์",
                ["noDataBelow"] = "ไม่พบข้อมูลใต้เซลล์ส่วนหัว",
                ["exportFailed"] = "การส่งออกล้มเหลว",
                ["devTag"] = "",
                ["exportCountAll"] = "ส่งออกจำนวน (ทั้งหมด)",
                ["exportCountFiltered"] = "ส่งออกจำนวน (กรองแล้ว)",
                ["sortBy"] = "การเรียงลำดับ",
                ["sortByCount"] = "ตามจำนวน",
                ["sortByTitle"] = "ตามหัวข้อ",
                ["sortOrder"] = "ลำดับ",
                ["sortDesc"] = "มากไปน้อย",
                ["sortAsc"] = "น้อยไปมาก",
                ["visualTools"] = "เครื่องมือภาพ",
                ["highlightToggle"] = "เน้นแถวและคอลัมน์",
                ["highlightColor"] = "สีเน้น"
            },
            [Language.id] = new Dictionary<string, string>
            {
                ["exportCount"] = "Ekspor Hitungan",
                ["exportCountScreentip"] = "Ekspor nilai unik dengan jumlah kemunculan",
                ["exportCountSupertip"] = "Pilih sel judul di kolom mana pun, lalu klik untuk mengekspor semua nilai unik dan jumlah kemunculannya ke lembar kerja baru, diurutkan menurun.",
                ["dataTools"] = "Alat Data",
                ["about"] = "Tentang",
                ["version"] = "FUWOA BETA",
                ["language"] = "Bahasa",
                ["langRestart"] = "Beberapa terjemahan perlu memulai ulang Excel.",
                ["count"] = "Jumlah",
                ["column"] = "Kolom",
                ["noExcelApp"] = "Tidak dapat mengakses instans aplikasi Excel.",
                ["selectOneCell"] = "Silakan pilih satu sel judul di kolom.",
                ["noDataBelow"] = "Tidak ada data di bawah sel judul.",
                ["exportFailed"] = "Ekspor gagal",
                ["devTag"] = "",
                ["exportCountAll"] = "Ekspor Hitungan (Semua)",
                ["exportCountFiltered"] = "Ekspor Hitungan (Difilter)",
                ["sortBy"] = "Urutkan",
                ["sortByCount"] = "Berdasarkan Jumlah",
                ["sortByTitle"] = "Berdasarkan Judul",
                ["sortOrder"] = "Urutan",
                ["sortDesc"] = "Menurun",
                ["sortAsc"] = "Menaik",
                ["visualTools"] = "Alat Visual",
                ["highlightToggle"] = "Sorot Baris & Kolom",
                ["highlightColor"] = "Warna Sorot"
            },
            [Language.bo] = new Dictionary<string, string>
            {
                ["exportCount"] = "ཨང་གྲངས་ཕྱིར་འདོན།",
                ["exportCountScreentip"] = "མ་འདྲ་བའི་གྲངས་ཀ་ཕྱིར་འདོན།",
                ["exportCountSupertip"] = "ཚོགས་གྲངས་གང་ཡིན་རུང་མགོ་ཡིག་གི་དྲ་ཐིག་གཅིག་འདེམས་རོགས། རྗེས་སུ་མི་འདྲ་བའི་གྲངས་ཐང་ཚང་མ་དང་དེ་དག་གི་འབྱུང་གྲངས་གཤམ་འོག་ནས་གོང་འོག་ཏུ་སྒྲིག་སྟེ་བྱང་བུ་གསར་པར་ཕྱིར་འདོན་བྱེད།",
                ["dataTools"] = "གཞི་གྲངས་ཡོ་བྱད།",
                ["about"] = "སྐོར།",
                ["version"] = "FUWOA BETA",
                ["language"] = "སྐད་ཡིག",
                ["langRestart"] = "སྒྱུར་བསྒྱུར་ཁ་ཤས་ Excel ཡང་བསྐྱར་བརྒྱུད་གཏོང་དགོས།",
                ["count"] = "ཨང་གྲངས།",
                ["column"] = "ཀ་རྟགས།",
                ["noExcelApp"] = "Excel མཉེན་སྒྲིག་གི་དཔེ་མཚན་ལེན་མི་ཐུབ།",
                ["selectOneCell"] = "ཀ་རྟགས་ཤིག་གི་མགོ་ཡིག་དྲ་ཐིག་གཅིག་རྐྱང་འདེམས་རོགས།",
                ["noDataBelow"] = "མགོ་ཡིག་དྲ་ཐིག་གི་འོག་ཏུ་གཞི་གྲངས་མི་འདུག",
                ["exportFailed"] = "ཕྱིར་འདོན་ལས་འཆར་མ་བྱུང་།",
                ["devTag"] = "",
                ["exportCountAll"] = "ཨང་གྲངས་ཕྱིར་འདོན། (ཡོངས།)",
                ["exportCountFiltered"] = "ཨང་གྲངས་ཕྱིར་འདོན། (འདེམས་པ།)",
                ["sortBy"] = "སྒྲིག་སྟངས།",
                ["sortByCount"] = "ཨང་གྲངས་ལྟར།",
                ["sortByTitle"] = "མགོ་ཡིག་ལྟར།",
                ["sortOrder"] = "གོ་རིམ།",
                ["sortDesc"] = "ཡས་མས།",
                ["sortAsc"] = "མས་ཡས།",
                ["visualTools"] = "མིག་སྣང་ཡོ་བྱད།",
                ["highlightToggle"] = "གྲལ་ཐིག་དང་ཀ་རྟགས་མངོན་གསལ།"
            },
            [Language.ug] = new Dictionary<string, string>
            {
                ["exportCount"] = "ساناق چىقىرىش",
                ["exportCountScreentip"] = "تەكرارلانمىغان قىممەتلەرنى سانى بىلەن چىقىرىش",
                ["exportCountSupertip"] = "ھەرقانداق ئىستوننىڭ ماۋزۇ كاتەكچىسىنى تاللاڭ، ئاندىن چېكىپ تەكرارلانمىغان بارلىق قىممەتلەر ۋە ئۇلارنىڭ پەيدا بولۇش سانىنى چۈشۈش تەرتىپى بويىچە يېڭى ۋاراققا چىقىرىڭ.",
                ["dataTools"] = "سانلىق مەلۇمات قوراللىرى",
                ["about"] = "ھەققىدە",
                ["version"] = "FUWOA BETA",
                ["language"] = "تىل",
                ["langRestart"] = "بەزى تەرجىمىلەر Excel نى قايتا باشلاشنى تەلەپ قىلىدۇ.",
                ["count"] = "سان",
                ["column"] = "ئىستون",
                ["noExcelApp"] = "Excel پروگرامما ئىنستانسىغا ئېرىشەلمىدى.",
                ["selectOneCell"] = "ئىستوندىكى بىر ماۋزۇ كاتەكچىنى تاللاڭ.",
                ["noDataBelow"] = "ماۋزۇ كاتەكچىنىڭ ئاستىدا سانلىق مەلۇمات يوق.",
                ["exportFailed"] = "ساناق چىقىرىش مەغلۇپ بولدى",
                ["devTag"] = "",
                ["exportCountAll"] = "ساناق چىقىرىش (ھەممىسى)",
                ["exportCountFiltered"] = "ساناق چىقىرىش (سۈزۈلگەن)",
                ["sortBy"] = "تەرتىپلەش",
                ["sortByCount"] = "سان بويىچە",
                ["sortByTitle"] = "ماۋزۇ بويىچە",
                ["sortOrder"] = "تەرتىپ",
                ["sortDesc"] = "چۈشۈش",
                ["sortAsc"] = "ئۆرلەش",
                ["visualTools"] = "كۆرۈنۈش قوراللىرى",
                ["highlightToggle"] = "قۇر ۋە ئىستوننى يورۇتۇش",
                ["highlightColor"] = "يورۇتۇش رەڭگى"
            },
            [Language.ja] = new Dictionary<string, string>
            {
                ["exportCount"] = "カウント出力",
                ["exportCountScreentip"] = "ユニーク値と出現回数を出力",
                ["exportCountSupertip"] = "任意の列のヘッダーセルを選択し、クリックするとその列のすべてのユニーク値と出現回数を降順で新しいシートに出力します。",
                ["dataTools"] = "データツール",
                ["about"] = "バージョン情報",
                ["version"] = "FUWOA BETA",
                ["language"] = "言語",
                ["langRestart"] = "一部の翻訳は Excel の再起動後に反映されます。",
                ["count"] = "件数",
                ["column"] = "列",
                ["noExcelApp"] = "Excel アプリケーションのインスタンスを取得できません。",
                ["selectOneCell"] = "列のヘッダーセルを1つ選択してください。",
                ["noDataBelow"] = "ヘッダーセルの下にデータがありません。",
                ["exportFailed"] = "カウント出力に失敗しました",
                ["devTag"] = "",
                ["exportCountAll"] = "カウント出力 (すべて)",
                ["exportCountFiltered"] = "カウント出力 (フィルター)",
                ["sortBy"] = "並べ替え",
                ["sortByCount"] = "数値順",
                ["sortByTitle"] = "タイトル順",
                ["sortOrder"] = "順序",
                ["sortDesc"] = "降順",
                ["sortAsc"] = "昇順",
                ["visualTools"] = "ビジュアルツール",
                ["highlightToggle"] = "行列ハイライト",
                ["highlightColor"] = "ハイライト色"
            },
        };

        public static Language Current
        {
            get
            {
                if (!_loaded) Load();
                return _current;
            }
        }

        private static void Load()
        {
            _loaded = true;
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegPath);
                string saved = key?.GetValue(LangKey) as string;
                _current = saved switch
                {
                    "zh-TW" => Language.zh_TW,
                    "en" => Language.en,
                    "de" => Language.de,
                    "fr" => Language.fr,
                    "ru" => Language.ru,
                    "vi" => Language.vi,
                    "th" => Language.th,
                    "id" => Language.id,
                    "bo" => Language.bo,
                    "ug" => Language.ug,
                    "ja" => Language.ja,
                    _ => Language.zh_CN
                };
            }
            catch
            {
                _current = Language.zh_CN;
            }
        }

        public static void SetLanguage(Language lang)
        {
            _current = lang;
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RegPath);
                key.SetValue(LangKey, lang switch
                {
                    Language.zh_TW => "zh-TW",
                    Language.en => "en",
                    Language.de => "de",
                    Language.fr => "fr",
                    Language.ru => "ru",
                    Language.vi => "vi",
                    Language.th => "th",
                    Language.id => "id",
                    Language.bo => "bo",
                    Language.ug => "ug",
                    Language.ja => "ja",
                    _ => "zh-CN"
                });
            }
            catch { }
        }

        public static string Get(string key)
        {
            if (!_loaded) Load();
            return Strings.TryGetValue(_current, out var dict) &&
                   dict.TryGetValue(key, out var value)
                ? value
                : key;
        }
    }
}
