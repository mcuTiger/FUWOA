# FUWOA

FUWOA is a practical utility add-in for Excel.

## Features

- **Single Column Count Export**: Select a header cell in any column and click to export all unique values and their occurrence counts in descending order to a new worksheet.

More features are in the works.

## Supported Languages

12 languages:

简体中文 · 繁體中文 · English · Deutsch · Français · Русский · Tiếng Việt · ไทย · 日本語 · Bahasa Indonesia · བོད་སྐད། · ئۇيغۇرچە

> Language must be switched manually from the Excel add-in ribbon. It does not follow the system language. Default is Simplified Chinese.

## System Requirements

- Windows 10 / 11 (x64)
- Microsoft Office 2016 / 2019 / 2021 / Microsoft 365 (desktop, x64)
- .NET Framework 4.8

## Build

```sh
msbuild FUWOA.sln /p:Configuration=Release /p:Platform=x64
```
