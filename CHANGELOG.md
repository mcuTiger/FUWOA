# 更新日志

FUWOA 所有重要变更均记录于此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### 新增
- **Markdown 表格导入** — 从 .md 文件或剪贴板解析 Markdown 表格，弹窗勾选后可多选，每个表格写入独立的新工作表；表头加粗、自动列宽。支持 12 语言。

## [1.0.2.0] - 2026-08-10

### 变更
- **安装程序：WiX MSI 改进** — COM 注册由手动 HKCR 条目切换为 `RegAsm /codebase` CustomAction（延迟执行），程序集版本和强名称变更时不再需要手动维护注册表。
- **安装程序：注册表路径统一** — Load behavior 键从 `HKCU\...\16.0\...` 移至 `HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn`，与插件自身的读写路径一致，支持每机安装。
- **安装程序：版本管理** — `ProductCode="*"` 每次构建自动生成新 GUID；`Version` 通过预处理器变量 `-dProductVersion` 传入。通过固定 `UpgradeCode` 处理 MajorUpgrade。
- **构建：自动化流水线** — 新增 `build_msi.bat`，集成 MSBuild（AddIn + Core）与 `wix build`（MSI 打包）为一步，输出至 `dist\`。
- 程序集版本升至 `1.0.2.0`。

### 新增
- `Launch` 条件检查 `VersionNT64` — 32 位 Windows 上阻止安装并显示明确提示。

### 移除
- 删除废弃产物：旧 MSI 文件（`FUWOA_V1.0.0.wxs`、`FUWOA_V1.0.0.msi`）、Inno Setup 脚本（`FUWOA_Setup.iss`）、调试日志（`install.log`、`COM.reg`）。
- 移除 `Product.wxs` 中脆弱的硬编码 `ComRegistry` 组件（硬编码 CLSID/ProgId/InprocServer32 条目）。

### 修复
- 修正 `Product.wxs` 中 `Fuwoa.Core.dll` 的源路径，从 `Fuwoa.AddIn\bin\` 改为 `Fuwoa.Core\bin\`。

---

## [1.0.0.0] - 2026-08-09

### 变更
- **移除 BETA 标识** — 关于页面和版本信息现显示 "V1.0.0"。
- 程序集版本设为 `1.0.0.0`。

### 新增
- 初始 WiX MSI 安装程序（`Product.wxs`），含手动 COM 注册表条目和基于 HKCU 的 load behavior。
- `Registration.cmd` 供开发者侧手动 COM 注册。

### 性能优化（自 BETA 阶段承继）
- **导出计数 / 按列拆分**：`.Text` 切换为 `.Value2` 配合批量数组读写（分块大小 20,000 行）。差异仅限于使用自定义数字格式的单元格。
- **导出计数**：`ExportCountService` 中新增稳定次排序键（`ThenBy`），确保输出顺序确定。
- **ExcelGuard**：长耗时操作包裹 `ScreenUpdating=false`、`Calculation=xlManual`、`EnableEvents=false` 保护。
- **高亮管理器**：条件格式范围通过 `GetCfRange` 限定而非整张工作表；`_formattedSheets` 注册表防止重复应用。`ColorImageCache` 去重高亮图标。
- **Connect.cs**：匿名事件处理委托替换为命名字段（`_sheetActivateHandler`、`_sheetBeforeDeleteHandler`），确保 `OnDisconnection` 中可正确取消订阅。
- 新增 `ColumnIndexToLetter` 辅助方法，将 1-based 列索引转换为 Excel 列字母。

### 新增
- 12 语言支持框架（`LanguageManager` 基于字典查找）：简体中文、繁體中文、English、Deutsch、Français、Русский、Tiếng Việt、ไทย、日本語、Bahasa Indonesia、བོད་སྐད།、ئۇيغۇرچە。
- 导出计数中新增百分比列选项。
- 导出计数支持按标签或计数、升序/降序排序。
- 导出计数中新增仅可见行模式（遵循当前筛选器）。

---

## [0.x.x] - BETA 阶段

### 新增
- **导出计数**：选中列标题单元格，该列所有唯一值及其出现次数导出至新工作表。
- **按列拆分**：将源表格按所选列的值分组拆分至多个工作表。
- **行列高亮**：高亮当前行与列，便于大型数据集中导航定位。
- 功能区 UI 集成（`IRibbonExtensibility`），注册为 COM 加载项（`IDTExtensibility2`）。
- 强名称签名（`Fuwoa.snk`）。
- 多语言架构，支持 12 种语言。
