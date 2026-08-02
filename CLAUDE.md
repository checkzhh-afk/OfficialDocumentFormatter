# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

本目录用于开发「公文格式化」WPS 宏工具：读取公文文档并按《公司公文格式主要规范》（未尽事项参照 GB/T 9704-2012）自动套用字体、字号、对齐、缩进、行距等格式。

## 命令

```powershell
# 语法检查（方案A 主本）
node --check "方案A_ribbonUI插件\OfficialDocumentFormat_0.01\main.js"

# 运行识别逻辑测试（在 测试\ 目录下执行）
cd 测试
node "公文格式化识别逻辑测试.js"

# 安装方案A 插件到本机 WPS（写入 %APPDATA%\kingsoft\wps\jsaddons，重启 WPS 后出现「公文格式」选项卡）
powershell -File "方案A_ribbonUI插件\install_wps_addin.ps1"

# 辅助工具：将 Markdown 文稿转为公文格式 docx（需 npm i docx）
node "辅助工具\md2docx_gongwen.js" 输入.md 输出.docx
```

测试脚本无单测框架，直接以 `node` 跑断言、全部 PASS 即通过；新增识别规则时须在该脚本追加用例。

## 架构

存在**两套并行的格式化方案**，共享同一套公文识别与格式规则，但触发/部署方式不同：

- **方案A（ribbonUI 插件）** — `方案A_ribbonUI插件/`
  - `OfficialDocumentFormat_0.01/main.js` 是**唯一主本**（编辑只改这里）；`ribbon.xml` 定义选项卡按钮，通过 `onAction` 绑定 `main.js` 里的全局函数（`FormatOfficialDocument` 全文格式化、`Apply*Format` 选区格式化）。
  - `install_wps_addin.ps1` 把这两个文件复制到 WPS jsaddons 目录并生成 `publish.xml`。
  - `main.js` 结构：`ensureOfficialDocumentFormatter()` 惰性构建 `OfficialDocumentFormatter` 对象（挂载 `CONST` 常量与所有识别函数）；识别函数（`isHeading`/`isYiShi`/`isDateText`/`isAttachmentNote` 等）纯逻辑、可被测试脚本沙箱调用；`GetOfficialDocumentFormatter()` 是测试入口。

- **方案B（控制台）** — `方案B_控制台/`
  - `公文格式化控制台WPS宏.js` 自包含完整格式化逻辑，通过 `Console*` 函数以「控制台文档 + 目标文档」模式操作：选目标 → 记录选区 → 套格式。为兼容信创 Linux WPS（无 VBA/COM/ActiveX，文件选择框可能不可用），提供 `ConsoleRunMenu` 序号菜单和 `ConsoleUseActiveDocument` 兜底。
  - 按钮名到函数名的映射见该目录 `公文格式化控制台使用说明.md` / `公文格式化控制台按钮模板.md`。

关键约束：脚本运行于 **WPS JS 宏**环境，不使用 VBA/COM/ActiveX；字体名不带 GB2312 后缀（`"仿宋"`/`"宋体"`/`"黑体"`/`"楷体"`）；中文字体与西文 Times New Roman 分开设置，且**必须先设中文字体再设西文字体**，否则 WPS 会用中文字体覆盖西文（见 `测试/公文格式化运行结果比对与修复说明.md`）。

## 测试与规则来源

- `测试/公文格式化识别逻辑测试.js` 用 Node `vm` 沙箱加载方案A 主本 `main.js`（相对路径 `../方案A_ribbonUI插件/...`），对 `_inspect.json`（从测试文档抽取的段落结构）跑识别断言。**修改主本路径会破坏此引用。**
- 格式规范权威来源：`规范依据/公文规范.txt`。需求与格式速查表：`方案A_ribbonUI插件/公文格式化WPS宏需求与测试说明.md`。已知需求歧义的处理约定记录在 `测试/公文格式化测试结果.md`（如「一是/二是」句式、多行标题短行、日期字符宽度）。
- `辅助工具/md2docx_gongwen.js` 的字体/字号/行距常量须与主本规则保持一致。

## 注意

- 方案A 只有一份主本（`main.js`），修改后测试会自动读到，无需再同步任何副本。
- 检查 docx 产生的临时文件统一用 `_tmp` 前缀，用完删除。
- `业务样例/` 与 `测试/*结果*.docx` 为生成产物，非任务要求勿改。
