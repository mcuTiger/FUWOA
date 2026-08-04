---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 71bb0250790dfa3e5617675e845c91af_ecb9d932903911f1bafa525400287e28
    ReservedCode1: PuAzv6Whti5L19FDD87oua0urtP25sjzqXlUlilmlxXBQIBJDWEbYox+UE2gEuWHV87gPiNcmWjLrWtW7rn/7OmkFBhUpS2sfXk9H7+oNaZd7j1nNQE55iBbj5gWejdzSbhz88H+7JiskUMdDK/fFrE61JZRiaE2R94RiVSWK32XjzuW4Zxhj1ilRXM=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 71bb0250790dfa3e5617675e845c91af_ecb9d932903911f1bafa525400287e28
    ReservedCode2: PuAzv6Whti5L19FDD87oua0urtP25sjzqXlUlilmlxXBQIBJDWEbYox+UE2gEuWHV87gPiNcmWjLrWtW7rn/7OmkFBhUpS2sfXk9H7+oNaZd7j1nNQE55iBbj5gWejdzSbhz88H+7JiskUMdDK/fFrE61JZRiaE2R94RiVSWK32XjzuW4Zxhj1ilRXM=
---

# FUWOA

Excel COM 加载项，选中列标题即可将唯一值计数导出到新工作表。

## 功能特性

- **列标题计数导出** — 支持单列或多列
- **排序方式** — 按数值或按标题
- **排序方向** — 升序或降序
- **筛选状态感知** — 筛选时标签自动变化
- **12 种语言** — 根据 Excel 界面语言自动切换
- **重名自动编号** — 工作表同名时自动添加序号

## 安装方式

### DEV 构建

```sh
git clone <仓库地址>
# 在 Visual Studio 中打开 FUWOA.sln
# 生成 → 配置: Debug, 平台: x64
# 注册: RegAsm.exe /codebase FUWOA.dll
```

### MSI 安装

下载最新 Release 中的 `.msi` 安装包，安装后以管理员身份运行 `Registration.cmd` 完成注册。

## 技术栈

.NET Framework 4.8 · C# · IDTExtensibility2 + IRibbonExtensibility · WiX Toolset

## 许可证

[GPL-3.0](LICENSE)
*（内容由AI生成，仅供参考）*
