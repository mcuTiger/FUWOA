# Changelog

All notable changes to FUWOA will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.2.0] - 2026-08-10

### Changed
- **Installer: WiX MSI improved** — COM registration switched from manual HKCR entries to `RegAsm /codebase` CustomAction (deferred), surviving assembly version and strong name changes.
- **Installer: registry path unified** — Load behavior key moved from `HKCU\...\16.0\...` to `HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn`, matching the plugin's own read/write path and enabling per-machine install.
- **Installer: version management** — `ProductCode="*"` auto-generates a new GUID per build; `Version` sourced from a preprocessor variable (`-dProductVersion`). MajorUpgrade handled by fixed `UpgradeCode`.
- **Build: automated pipeline** — New `build_msi.bat` orchestrates MSBuild (AddIn + Core) and `wix build` (MSI packaging) in one step, outputting to `dist\`.
- Assembly version bumped to `1.0.2.0`.

### Added
- `Launch` condition checking `VersionNT64` — installation blocked on 32-bit Windows with a clear message.

### Removed
- Deleted stale artifacts: old MSI files (`FUWOA_V1.0.0.wxs`, `FUWOA_V1.0.0.msi`), Inno Setup script (`FUWOA_Setup.iss`), debug logs (`install.log`, `COM.reg`).
- Removed fragile manual `ComRegistry` component from `Product.wxs` (hardcoded CLSID/ProgId/InprocServer32 entries).

### Fixed
- `Fuwoa.Core.dll` source path in `Product.wxs` corrected from `Fuwoa.AddIn\bin\` to `Fuwoa.Core\bin\`.

---

## [1.0.0.0] - 2026-08-09

### Changed
- **BETA label removed** — About page and version info now display "V1.0.0".
- Assembly version set to `1.0.0.0`.

### Added
- Initial WiX MSI installer (`Product.wxs`) with manual COM registry entries and HKCU-based load behavior.
- `Registration.cmd` for developer-side manual COM registration.

### Performance improvements (carried from BETA)
- **Export Count / Split by Column**: switched from `.Text` to `.Value2` with batch array read/write (chunk size 20,000 rows). Differences are limited to cells with custom number formats.
- **Export Count**: added stable secondary sort key (`ThenBy`) in `ExportCountService` to ensure deterministic output ordering.
- **ExcelGuard**: wraps long-running operations with `ScreenUpdating=false`, `Calculation=xlManual`, `EnableEvents=false` protection.
- **Highlight Manager**: conditional format range now scoped via `GetCfRange` instead of entire worksheet; `_formattedSheets` registry prevents redundant reapplication. `ColorImageCache` deduplicates highlight icons.
- **Connect.cs**: anonymous event handler delegates replaced with named fields (`_sheetActivateHandler`, `_sheetBeforeDeleteHandler`) to enable proper unsubscription in `OnDisconnection`.
- Added `ColumnIndexToLetter` helper for converting 1-based column indices to Excel column letters.

### Added
- 12-language support framework (`LanguageManager` with dictionary-based lookups): 简体中文, 繁體中文, English, Deutsch, Français, Русский, Tiếng Việt, ไทย, 日本語, Bahasa Indonesia, བོད་སྐད།, ئۇيغۇرچە.
- Percentage column option in Export Count.
- Sorting by label or count, ascending/descending in Export Count.
- Visible-rows-only mode in Export Count (respects active filter).

---

## [0.x.x] - BETA Phase

### Added
- **Export Count**: select a column header cell, and all unique values below it plus their occurrence counts are exported to a new worksheet.
- **Split by Column**: split a source table into multiple worksheets grouped by a chosen column's values.
- **Row & Column Highlighting**: highlight the current row and column for navigation in large datasets.
- Ribbon UI integration (`IRibbonExtensibility`), registered as a COM Add-in (`IDTExtensibility2`).
- Strong-name signing (`Fuwoa.snk`).
- Multi-language architecture with 12 supported languages.
