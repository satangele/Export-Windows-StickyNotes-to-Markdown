# Export-Windows-StickyNotes-to-Markdown

将 Windows 便签（Sticky Notes）的**全部文字与图片附件**无损导出为 Markdown 文件的 PowerShell 工具。

## 功能特性

- 从 `plum.sqlite` 读取全部便签，逐条导出为独立 `.md` 文件
- 精准匹配 `Media` 表与 `LocalState\media`，把图片复制到 `attachments/`，并以 Markdown 图片语法置顶展示
- 自动生成 YAML Frontmatter（`sticky_id`、`created`、`updated`、`title`、`color`、`tags`）
- 清洗便签富文本标记（`\b`、`\i`、`\ul`、`\strike`、`\l`、`\id` 等）并转换为 Markdown
- 严格剥离 `\id=UUID`，避免 UUID 混入标题与文件名
- 可选导出已删除便签（`-IncludeDeleted`）
- 同步文件创建 / 修改时间，保留原始时间线
- 自动定位 `sqlite3.exe`（本地 `sqlite\` 目录 → PATH → 手动输入）
- 导出前复制数据库快照到临时目录，避免锁库

## 环境要求

- Windows 10 / 11
- Windows PowerShell 5.1 或 PowerShell 7+
- `sqlite3.exe`（SQLite 命令行工具）
- 已安装并至少启动过一次 Windows 便签（Sticky Notes）

## 安装

```powershell
git clone https://github.com/satangele/Export-Windows-StickyNotes-to-Markdown.git
cd Export-Windows-StickyNotes-to-Markdown
```

准备 `sqlite3.exe`，任选其一：

1. 放到脚本同级的 `sqlite\sqlite3.exe`（推荐）
2. 加入系统 `PATH`
3. 运行时通过 `-SqlitePath` 指定

## 使用

交互式运行：

```powershell
.\Export-StickyNotes.ps1
```

指定导出目录：

```powershell
.\Export-StickyNotes.ps1 -OutputPath "D:\StickyNotes_Export"
```

同时导出已删除的便签：

```powershell
.\Export-StickyNotes.ps1 -IncludeDeleted
```

指定 sqlite3 路径与附件目录名：

```powershell
.\Export-StickyNotes.ps1 -SqlitePath "C:\tools\sqlite3.exe" -MediaDirName "images"
```

若执行策略受限：

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-StickyNotes.ps1
```

## 参数

| 参数 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `-OutputPath` | `string` | 否 | 导出目录，默认 `.\StickyNotes_Export`，也可交互输入 |
| `-MediaDirName` | `string` | 否 | 图片子目录名，默认 `attachments` |
| `-IncludeDeleted` | `switch` | 否 | 同时导出已删除便签 |
| `-SqlitePath` | `string` | 否 | 指定 `sqlite3.exe` 的完整路径 |

## 输出结构

```text
StickyNotes_Export/
├── 便签标题.md
├── 另一个标题.md
└── attachments/
    ├── media_<noteId>_1.png
    └── ...
```

每个 Markdown 文件的结构：

```markdown
---
sticky_id: "<note-id>"
created: "2026-01-01 09:00:00"
title: "便签标题"
color: "Yellow"
updated: "2026-01-02 10:30:00"
tags:
  - "sticky-notes"
  - "color/yellow"
---

![图片说明](attachments/media_<noteId>_1.png)

正文内容……
```

## 便签数据库位置

```text
%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite
```

## 注意事项

- 导出前建议关闭便签应用；脚本会先复制 `plum.sqlite*` 快照再读取，降低锁库风险。
- 图片文件名带便签 ID，避免不同便签的同名图片互相覆盖。
- 本工具只读取本地数据，不会上传任何内容。
- 导出结果可能包含个人隐私，请勿将 `StickyNotes_Export/` 提交到公开仓库（已在 `.gitignore` 中忽略）。
