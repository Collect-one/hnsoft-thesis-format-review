# hnsoft-thesis-format-review

湖南软件职业技术大学毕业论文/毕业设计 DOCX 的格式、结构与内容一致性审查 Skill。

## 功能

- 检查标题、正文、图题、代码的字体、字号、行距和缩进
- 按总体功能模块图检查角色名称、角色顺序和章节结构；角色型第4章的三级功能标题按项目实际功能核对，不要求与图节点逐项对应，允许不同角色出现同名功能
- 检查角色型标题、图题、表题和简易图中的角色名称是否统一
- 检查目录缓存页码与实际分页是否一致
- 检查图题、正文引用、图片复用和编号连续性
- 检查三线表宽度、边线、表题和行结构
- 检查参考文献编号、年份、标点和著录项
- 检查摘要、关键词、英文摘要和结论/参考文献/致谢标题
- 导出 PDF 并渲染 PNG，便于逐页目检
- 默认只读；只有用户明确授权后才做最小修改
- 版面美化、页尾空白和代码框调整优先交给配套的 `hnsoft-thesis-layout-polish`；未安装该 Skill 时，本 Skill 仍完成格式、结构和内容审查，只列出版面问题，不自动安装配套 Skill

## 快速安装

把下面这句话发给你的 Agent：

```text
帮我安装这个 skill：https://github.com/Collect-one/hnsoft-thesis-format-review
```

<details>
<summary><strong>Agent 不支持直接安装时</strong></summary>

不同 Agent 的 skills 目录不同；下面以 Codex 为例。

手动克隆：

```powershell
git clone https://github.com/Collect-one/hnsoft-thesis-format-review.git "$env:USERPROFILE\.codex\skills\hnsoft-thesis-format-review"
```

如果设置了 `CODEX_HOME`，目标路径为：

```text
%CODEX_HOME%\skills\hnsoft-thesis-format-review
```

更新已安装版本：

```powershell
cd "$env:USERPROFILE\.codex\skills\hnsoft-thesis-format-review"
git pull
```

</details>

## 装好之后这样用

```text
使用 $hnsoft-thesis-format-review，检查这篇毕业论文的格式和前后矛盾。
```

## 目录结构

```text
hnsoft-thesis-format-review/
├─ SKILL.md
├─ agents/openai.yaml
├─ scripts/
│  ├─ audit_all.py
│  ├─ inspect_docx.py
│  ├─ check_figures.py
│  ├─ check_tables.py
│  ├─ check_toc.py
│  ├─ check_refs.py
│  ├─ check_front_matter.py
│  ├─ docx_common.py
│  ├─ export_pdf.ps1
│  └─ render_pages.py
└─ references/
   ├─ hunanruan-2026-format.md
   ├─ review-checklist.md
   ├─ thesis-structure-profile.md
   └─ report-template.md
```

## 依赖

- Python 3.10+（审计脚本跨平台）
- PyMuPDF（`pip install pymupdf`），仅 `render_pages.py` 需要
- Windows + Microsoft Word，仅 `export_pdf.ps1` 需要；macOS/Linux 或其他环境可先手动导出 PDF，再用 `render_pages.py`
- PowerShell，仅 `export_pdf.ps1` 需要

`export_pdf.ps1` 会先检查 Word 的 `~$` 锁文件，文件被 Word 占用时会拒绝导出，避免影响正在编辑的文档。

## 说明

- 本 Skill 主要面向湖南软件职业技术大学 **2026年10月版**维普格式，适用于 **2027届**毕业生；其他学校使用时，请替换 `references/hunanruan-2026-format.md`。
- 学校最新范本和格式要求优先于本仓库的规则摘要。
- `格式要求.txt` 是可选输入示例，不是硬依赖；规则已内置于 `references/hunanruan-2026-format.md`。
- 默认只读；修改论文前需要用户明确授权，并先备份、再重新导出 PDF 核对分页。
- 开放式版面美化不属于本 Skill。已安装 `hnsoft-thesis-layout-polish` 时使用该 Skill 并在完成后复审；未安装时保留版面问题清单，不声称已完成美化。
- 本仓库不包含学校范本原文件，只包含规则摘要和审查脚本。
- `agents/openai.yaml` 是 Codex 的可选 UI 元数据；其他 Agent 可以忽略。

## 配套版面优化 Skill

本 Skill 负责格式、结构、内容一致性和参考文献审查，不负责整篇论文的版面美化。

如需处理页尾空白、图表分页、续表、图片尺寸、代码框和目录同步，可另行安装：

- 仓库：[hnsoft-thesis-layout-polish](https://github.com/Collect-one/hnsoft-thesis-layout-polish)

也可以直接对 Agent 说：

```text
帮我安装这个 skill：
https://github.com/Collect-one/hnsoft-thesis-layout-polish
```

两个 Skill 可以独立安装。未安装版面优化 Skill 时，本 Skill 仍完成格式、结构和内容审查，只输出版面问题清单，不会自动下载或安装另一个 Skill。

## 免责声明

本 Skill 的规则摘要基于 2026 年 10 月学校范本整理，适用于 2027 届毕业生，仅供审查辅助使用。提交前请以学校最新官方格式要求和导师意见为准。

