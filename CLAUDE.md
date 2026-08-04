# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

本目录用于开发「公文格式化」WPS 宏工具：读取公文文档并按《公司公文格式主要规范》（未尽事项参照 GB/T 9704-2012）自动套用字体、字号、对齐、缩进、行距等格式。

## 命令

```powershell
# 语法检查（方案A 主本）
node --check "方案A_ribbonUI插件\OfficialDocumentFormat_0.01\main.js"

# 语法检查（方案B 主本）
node --check "方案B_控制台\公文格式化控制台WPS宏.js"

# 运行完整回归测试（编码、ES5、宏入口、A/B一致性、冷启动、页码和识别逻辑）
cd 测试
node "公文格式化识别逻辑测试.js"

# 运行方案B交互模拟与DOCM文档包结构测试
node "方案B控制台交互与文档结构测试.js"

# 重新生成内嵌最新宏源码和23个按钮的方案B控制台
cd ..
node "方案B_控制台\生成公文格式控制台文档.js"

# 安装方案A 插件到本机 WPS（写入 %APPDATA%\kingsoft\wps\jsaddons，重启 WPS 后出现「公文格式」选项卡）
powershell -File "方案A_ribbonUI插件\install_wps_addin.ps1"

# 辅助工具：将 Markdown 文稿转为公文格式 docx（需 npm i docx）
node "辅助工具\md2docx_gongwen.js" 输入.md 输出.docx
```

测试脚本无单测框架，直接以 `node` 跑断言、全部 PASS 即通过；新增识别规则时须在该脚本追加用例。

## 架构

存在**两套并行的格式化方案**，共享同一套公文识别与格式规则，但触发/部署方式不同：

- **方案A（ribbonUI 插件）** — `方案A_ribbonUI插件/`
  - `OfficialDocumentFormat_0.01/main.js` 是方案 A 的**唯一部署主本**；`ribbon.xml` 定义选项卡按钮，通过 `onAction` 绑定 `main.js` 里的全局函数（`FormatOfficialDocument` 全文格式化、`Apply*Format` 选区格式化）。
  - `install_wps_addin.ps1` 把这两个文件复制到 WPS jsaddons 目录并生成 `publish.xml`。
  - `main.js` 结构：`ensureOfficialDocumentFormatter()` 惰性构建 `OfficialDocumentFormatter` 对象（挂载 `CONST` 常量与所有识别函数）；识别函数（`isHeading`/`isYiShi`/`isDateText`/`isAttachmentNote` 等）纯逻辑、可被测试脚本沙箱调用；`GetOfficialDocumentFormatter()` 是测试入口。

- **方案B（控制台）** — `方案B_控制台/`
  - `公文格式控制台.docm` 是正式运行入口，内嵌 WPS JSA 源码、JDE 关系和 23 个复合 `MACROBUTTON` 域。每项显示为“双击：操作名称”，用户双击完整动作文字执行。
  - `公文格式化控制台WPS宏.js` 是唯一源码主本；`生成公文格式控制台文档.js` 负责把最新源码和按钮重新写入 `DOCM`。
  - 同目录全文格式化支持 `.doc/.docx/.docm/.wps/.wpt`：控制台自身和临时文件不得进入候选；唯一候选自动选择；多个候选优先自动打开原生文件选择框，用户可用上下方向键选择并按 Enter；只有 `FileDialog` 不可用时才回退到输入序号。全文格式化完成后保持目标文档为活动文档。
  - 控制台按钮使用上/左高光、下/右阴影的立体边框；WPS `MACROBUTTON` 的第一次单击可能只选中字段，双击完整动作文字执行。
  - 局部按钮自动读取其他已打开文档窗口中保留的选区；多个文档时提示选择，同一文档多窗口时优先使用非折叠选区。
  - `ConsoleRunMenu` 和单独导入 JS 是兼容兜底，不是正常使用 `DOCM` 的前置步骤。
  - 用户操作见项目根目录 `使用说明.md` 及该目录 `安装与使用说明.md`。旧 `公文格式控制台.docx`、`原型_*` 和 `历史参考/` 不得作为正式入口。

关键约束：脚本运行于 **WPS JS 宏**环境，不使用 VBA/COM/ActiveX；字体名不带 GB2312 后缀（`"仿宋"`/`"宋体"`/`"黑体"`/`"楷体"`）；中文字体与西文 Times New Roman 分开设置，且**必须先设中文字体再设西文字体**，否则 WPS 会用中文字体覆盖西文（见 `测试/公文格式化运行结果比对与修复说明.md`）。

两份正式宏必须保持以下源码约束：

- `/* SHARED_CORE_START */` 到 `/* SHARED_CORE_END */` 之间是完全相同的共享核心；修改识别、格式或页码逻辑时必须同步更新 A/B。
- 正式宏文件使用纯 7-bit ASCII 和 CRLF 换行，所有中文及其他非 ASCII 字符写成 `\uXXXX`，避免 WPS 宏编辑器按错误编码打开后破坏源码。
- 只使用 ES5 语法，不使用 `let`、`const`、箭头函数、模板字符串、可选链、空值合并、`async/await`。
- `测试/公文格式化识别逻辑测试.js` 会强制校验以上约束以及 Ribbon/控制台入口完整性。

## 测试与规则来源

- `测试/公文格式化识别逻辑测试.js` 用 Node `vm` 沙箱同时加载方案A主本和方案B控制台脚本，校验源码编码、ES5 兼容、共享核心、宏入口、冷启动局部格式和页码处理，再对 `_inspect.json`（从测试文档抽取的段落结构）及内联边界场景跑识别断言。**修改两套主本路径会破坏此引用。**
- `测试/方案B控制台交互与文档结构测试.js` 模拟同目录单文件、多文件原生选择框优先、无 `FileDialog` 时序号回退、多个打开文档和多窗口选区，并校验 `DOCM` 的宏内容类型、JDE 关系、内嵌源码、23组复合宏字段和可见动作文字。WPS对 `w:fldSimple` 的缓存文字显示不稳定，控制台不得恢复为简单域。
- 格式规范权威来源：`规范依据/公文规范.txt`。需求与格式速查表：`方案A_ribbonUI插件/公文格式化WPS宏需求与测试说明.md`。已知需求歧义的处理约定记录在 `测试/公文格式化测试结果.md`（如「一是/二是」句式、多行标题短行、日期字符宽度）。
- `辅助工具/md2docx_gongwen.js` 的字体/字号/行距常量须与主本规则保持一致。

## 注意

- 方案A 只有一份部署主本（`main.js`），但其共享核心必须与方案B标记区间完全一致。
- 方案B 只有一份源码主本（`方案B_控制台/公文格式化控制台WPS宏.js`）；每次修改后必须规范化为 CRLF、重新生成 `公文格式控制台.docm` 并运行两套测试。原 `脚本/` 目录已经合并，不应重新创建。
- 格式化会修改打开的文档，但不自动保存。
- 检查 docx 产生的临时文件统一用 `_tmp` 前缀，用完删除。
- `业务样例/` 与 `测试/*结果*.docx` 为生成产物，非任务要求勿改。
