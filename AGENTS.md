# FUWOA 开发指导（AI AGENT 使用）

## 0. 项目一句话

Excel 桌面版 COM 加载项（.NET Framework 4.7.2 / x64 / C# 8，旧式 csproj）。功能区 Ribbon + 数据工具（导出计数、导出分类、百分比列、按列拆分）+ 行列高亮 + 12 种语言。三项目结构：`Fuwoa.Core`（纯计算）、`Fuwoa.AddIn`（COM 入口与命令）、`Fuwoa.Installer`（WiX MSI，经 `build_msi.bat` 出包）。

## 1. 两套构建：DEV 与 RELEASE（最重要的背景，先读这节）

- **DEV = Debug|x64**（`DefineConstants=DEBUG;TRACE`）：功能区显示 DEVTAG 标签（zh_CN “当前是 Tiger 的开发测试版本”），开发测试专用。
- **RELEASE = Release|x64**（`DefineConstants=TRACE;RELEASE`）：打包上传 GitHub 用，隐藏 DEVTAG。
- **两套构建的唯一差异就是 devtag**。行列高亮（`HighlightManager` + 功能区“视觉工具”组）自 v1.0.2.0 起已是正式功能，DEV/RELEASE 都包含。
- 门控机制：`Connect.cs` 中 `#if !RELEASE` 目前**只**包住 devtag 标签一处。以后新增“仅开发版”内容时，集中用 `#if !RELEASE` 包住，不要散落条件编译；能做成运行时开关的优先用开关。
- **铁律：任何任务完成前，必须同时编译 Debug|x64 与 Release|x64 两个配置。** 只编 Debug 会让 Release 路径的编译错误/行为差异延迟到打包时才暴露。

## 2. 构建与出包命令

```powershell
# 每次改动后必跑（验证）
msbuild Fuwoa.sln /t:Build /p:Configuration=Debug /p:Platform=x64
msbuild Fuwoa.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

- `Fuwoa.Installer` 是 SDK 风格 WiX 项目（`WixToolset.Sdk`），随 sln 构建时需要能解析该 SDK（联网或本地 NuGet 缓存）；离线环境请直接用下面的出包脚本。
- 出包（Windows，需 WiX v7；首次在本机执行 `wix eula accept wix7` 接受 OSMF EULA）：
  ```powershell
  build_msi.bat    # 产物: dist\FUWOA_<版本>_x64.msi
  ```
- `build.bat`（Release 构建 + RegAsm COM 注册，需管理员）仅开发机用，**不作为验证手段**。

## 3. 目录结构与可改范围

| 路径 | 职责 | Agent 可改 |
|---|---|---|
| `Fuwoa.Core\ExportCount\ExportCountService.cs` | 分组计数/排序，纯逻辑、无 COM | ✅ 可改 |
| `Fuwoa.AddIn\Connect.cs` | COM 入口、功能区 XML、回调、设置持久化、devtag 门控 | ✅ 可改 |
| `Fuwoa.AddIn\Commands\*.cs` | 两个命令的 Excel 交互逻辑 | ✅ 可改 |
| `Fuwoa.AddIn\LanguageManager.cs` | 12 语言文案字典 | ✅ 可改（新增键必须补齐 12 语言；含 `version` 标签）|
| `Fuwoa.AddIn\HighlightManager.cs` | 行列高亮（正式功能，Release 也包含）| ✅ 可改 |
| `Fuwoa.AddIn\*.csproj` | 工程配置 | ⚠️ 谨慎：新增 .cs 必须手工登记 `<Compile>` 项 |
| `Fuwoa.Installer\*`、`build_msi.bat` | WiX 安装器与出包脚本 | ❌ 默认禁止（除非任务明确要求）|
| `build.bat`、`uninstall.bat`、`README*`、`LICENSE`、`Properties\AssemblyInfo.cs` | 构建脚本与文档 | ❌ 默认禁止（版本号同步除外，见 4.4）|
| `bin\`、`obj\`、`.vs\`、`dist\`、`*.msi`、`*.log`、`*.reg` | 产物 | 不提交、不修改 |

## 4. 硬性约束（违反即返工）

1. **不引入 NuGet/第三方依赖**：仅 .NET Framework 4.7.2 + Excel Interop。
2. **不改业务语义**：输出内容、表名/重名规则、筛选统计范围、多语言结构保持现状。
3. **语言键新增 → 12 种语言字典全部补齐**（缺键会在 UI 回退显示 key 本身）。
4. **版本号三处同步**：改版本时必须同时更新 `LanguageManager.cs` 的 `version` 键、`Properties\AssemblyInfo.cs` 的 `AssemblyVersion/FileVersion`、`build_msi.bat` 的 `VERSION=`（当前统一为 1.0.2.0）。
5. **不提交产物**：`bin/ obj/ .vs/ dist/`、MSI、日志、`.reg` 等一律不提交。
6. **保持代码风格**：C# 8、中文注释、现有命名与结构；不新增版权头。
7. **不擅自重构无关代码**：只改任务要求的范围。

## 5. 完成定义（Definition of Done）

任务“完成”必须全部满足：

- [ ] Debug|x64 编译通过
- [ ] Release|x64 编译通过
- [ ] 涉及 `Fuwoa.Core` 时：用临时脚本或测试验证分组/排序结果正确，验证后清理临时文件
- [ ] 涉及文案时：12 种语言补齐
- [ ] 边界自查：空值、筛选态、单行/单列、重名、31 字符表名、合并单元格、异常路径
- [ ] 汇报中列出：改动文件清单、关键决策、**行为差异**（如 `.Text→Value2`、排序并列顺序）、未验证项
- [ ] 有功能/行为变化时，同步更新 CHANGELOG.md（无对应版本条目则新建，格式照旧）
- [ ] 未触碰第 4 节禁止项

## 6. 任务执行流程（Agent 必须按此顺序）

1. **先读后改**：通读相关文件；若任务附带报告（如 `FUWOA_优化改造报告.md`、`FUWOA_安装程序改造报告.md`），以报告为准逐条实施；用 `git log`/`git blame` 了解代码背景。
2. **给计划**：小任务可直接动手；中大型任务先输出改动清单（文件 + 每处改什么 + 风险），再实施。
3. **小步改**：一次只改一个逻辑点；新增 `.cs` 记得登记进 csproj 的 `<Compile>`。
4. **双编译验证**：Debug + Release 都必须通过（见第 2 节）。
5. **自查边界**：参考第 5 节清单。
6. **汇报**：改动文件清单、关键决策、行为差异、未完成/未验证项。

## 7. 已知技术要点（避免踩坑）

- **Excel Interop 性能**：禁止逐格 `Cells[r,c].Value2/.Text` 大循环；用 `Range.Value2` 整块读、`object[,]` 整块写。`Value2` 对**单格**返回标量、对多格返回 1 基二维数组，两种都要处理。
- `SpecialCells(xlCellTypeVisible)` 无可视单元格时抛 1004，必须 try/catch 并给出友好提示。
- 批量操作建议包裹 `ScreenUpdating=false` + `Calculation=manual` + `EnableEvents=false`，`finally` 恢复，避免异常后 Excel 留在脏状态。
- **注册表**：用户设置存 `HKCU\SOFTWARE\Microsoft\Office\Excel\Addins\Fuwoa.AddIn`（**无 16.0 路径**）；高亮颜色存 `HKCU\SOFTWARE\FUWOA`。MSI 注册键在 HKLM（机器级）。
- 功能区 XML 由 `GetCustomUI` 动态生成；控件 id 改动要同步 `InvalidateControl(...)` 里的字符串。
- **条件编译边界**：`#if !RELEASE` 目前只包住 devtag；`HighlightManager` 已是双构建共有代码。跨门控的公共逻辑放门外。
- 新工作表名：Excel 限制 31 字符、禁 `[ ] / * ? : \`；重名加 ` (n)` 后缀，比较用 `OrdinalIgnoreCase`。

## 8. Git 规范

- **分支**：默认在 `main`（trunk-based），不建长命分支；发布 = 打 tag。
- **提交**：一次任务一个提交；信息格式 `类型: 简述`（如 `perf: 导出计数改为批量读写`、`fix: 分类表激活顺序`、`feat: 新增xx`、`docs: ...`），并注明行为差异。
- **提交前自查**：`git status` 确认没有 `bin/ obj/ .vs/ dist/ *.msi *.log` 混入。
- **版本 tag**：`v<主>.<次>.<修订>`（如 `v1.0.2`），与 `AssemblyVersion`、MSI 版本对应。
- **发版前**：先扫描 GitHub Issues，把待处理用户反馈整理成任务清单；已修复内容记入 CHANGELOG 对应版本。

## 9. 参考文档

- `C:\Users\Kevin\Documents\工作\FUWOA_优化改造报告.md` —— 性能与稳定性改造任务书（若任务相关，逐条实施）
- `C:\Users\Kevin\Documents\工作\FUWOA_安装程序改造报告.md` —— 安装程序改造任务书（仅当任务涉及安装器时）

## 附录 A：任务描述模板（给任务发起人复用）

> **背景**：（问题或目标，一句话）
> **范围**：（明确到文件/功能，如“只改 ExportCountCommand”）
> **验收**：（可检查的标准，如“10 万行导出计数 < 3 秒”“Debug/Release 双编译通过”）
> **必须执行**：`msbuild Fuwoa.sln /t:Build /p:Configuration=Debug /p:Platform=x64` 与 Release 同款命令
> **禁止**：（如有，写清，如“不动安装器”）

## 附录 B：常见任务类型速查

| 任务类型 | Agent 要点 |
|---|---|
| 性能优化 | 批量数组读写 + Excel 保护上下文 + 双编译 |
| 修 Bug | 先复现/读日志 → 定位 → 最小改动 → 边界自查 |
| 新功能 | 遵循现有结构；功能区按钮在 `GetCustomUI` 加；文案 12 语言 |
| 翻译调整 | 只改 `LanguageManager.cs` 字典，键名不变；注意 `version` 键与发版版本一致 |
| 打包/安装器 | 先读 `FUWOA_安装程序改造报告.md`，默认需另行授权 |
