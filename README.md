---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 71bb0250790dfa3e5617675e845c91af_ea897aa6903911f1bafa525400287e28
    ReservedCode1: R1SNZ/WqJjkYcv6jlShm/W5rzjGQV8pprnp/z0IgDxsl4bav94rkNettbssGOChzcm2oRIgD6xNtLU/G6CF/Em4ysg5/GzWT9oyLyTliaV92mgzRsTm3EGRC3YJwcYPY0hu4nMZxAuFajYqBNDo1lKj3iLXXSfTjpdE8p4sr1MP87jhAzvmrvzy+tiw=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 71bb0250790dfa3e5617675e845c91af_ea897aa6903911f1bafa525400287e28
    ReservedCode2: R1SNZ/WqJjkYcv6jlShm/W5rzjGQV8pprnp/z0IgDxsl4bav94rkNettbssGOChzcm2oRIgD6xNtLU/G6CF/Em4ysg5/GzWT9oyLyTliaV92mgzRsTm3EGRC3YJwcYPY0hu4nMZxAuFajYqBNDo1lKj3iLXXSfTjpdE8p4sr1MP87jhAzvmrvzy+tiw=
---

# FUWOA

Excel COM Add-in that exports column header counts to a new worksheet. Select column headers, get unique value frequency counts instantly.

## Features

- **Column Count Export** — single or multiple columns
- **Sort** — by value or by header label
- **Direction** — ascending or descending
- **Filter-Aware** — labels update automatically when filters are applied
- **12 Languages** — auto-detected from Excel UI language
- **Auto-Numbering** — duplicate sheet names get sequential suffixes

## Installation

### DEV Build

```sh
git clone <repo-url>
# Open FUWOA.sln in Visual Studio
# Build → Configuration: Debug, Platform: x64
# Register: RegAsm.exe /codebase FUWOA.dll
```

### MSI Install

Download the latest release `.msi`, install, then run `Registration.cmd` as Administrator to register the add-in with Excel.

## Tech Stack

.NET Framework 4.8 · C# · IDTExtensibility2 + IRibbonExtensibility · WiX Toolset

## License

[GPL-3.0](LICENSE)
*（内容由AI生成，仅供参考）*
