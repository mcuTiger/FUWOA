# FUWOA

FUWOA 是一款适用于 Excel 的实用小工具。

## 功能

- **单列计数导出**：选中某一列的标题单元格，将下方所有唯一值及其出现次数按降序导出到新工作表。

更多功能正在构思中。

## 支持的语言

12 种语言：

简体中文 · 繁體中文 · English · Deutsch · Français · Русский · Tiếng Việt · ไทย · 日本語 · Bahasa Indonesia · བོད་སྐད། · ئۇيغۇرچە

> 语言需在 Excel 插件功能区中手动切换，不会跟随系统语言变化。默认语言为简体中文。

## 系统要求

- Windows 10 / 11（x64）
- Microsoft Office 2016 / 2019 / 2021 / Microsoft 365（桌面版，x64）
- .NET Framework 4.8

## 构建

```sh
msbuild FUWOA.sln /p:Configuration=Release /p:Platform=x64
```
