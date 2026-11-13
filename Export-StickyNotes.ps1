<#
.SYNOPSIS
    Windows 便签 (Sticky Notes) 全量数据与磁盘媒体无损导出工具 (Schema 最终修复版)
.DESCRIPTION
    1. 彻底解决标题清洗中 \i 误杀 \id 导致的文件名包含 UUID 乱码问题；
    2. 精准匹配 Media 表中 LocalFileRelativePath 字段，从 LocalState\media 目录无损提取图片；
    3. 图片悬浮置顶展示（放置于 Frontmatter 与正文之间）；
    4. 规范生成 YAML Frontmatter，并同步文件物理时间戳。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [string]$MediaDirName = "attachments",

    [switch]$IncludeDeleted,

    [string]$SqlitePath
)

# ----------------- 1. 锁定基准工作目录 -----------------
$scriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
} else {
    (Get-Location).Path
}
Set-Location -Path $scriptDir

# 定位 sqlite3.exe
function Resolve-SqliteExe([string]$customPath, [string]$baseDir) {
    if (-not [string]::IsNullOrWhiteSpace($customPath) -and (Test-Path $customPath)) {
        return (Resolve-Path $customPath).Path
    }
    $localSqlite = Join-Path $baseDir "sqlite\sqlite3.exe"
    if (Test-Path $localSqlite) { return (Resolve-Path $localSqlite).Path }

    $cmd = Get-Command sqlite3.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    return $null
}

$sqliteExe = Resolve-SqliteExe $SqlitePath $scriptDir
if (-not $sqliteExe) {
    if ([Environment]::UserInteractive) {
        Write-Host "[提示] 未在 .\sqlite\ 目录下检测到 sqlite3.exe。" -ForegroundColor Yellow
        $userPath = Read-Host "请输入 sqlite3.exe 完整路径"
        if (Test-Path $userPath) { $sqliteExe = (Resolve-Path $userPath).Path }
    }
    if (-not $sqliteExe) { throw "未找到 sqlite3.exe，请将该工具放入脚本下的 sqlite 目录中。" }
}

# 数据库路径与 LocalState 根目录
$plumPath = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
$localStateDir = [System.IO.Path]::GetDirectoryName($plumPath)
if (!(Test-Path $plumPath)) {
    throw "未找到便签数据库文件: $plumPath。请先在 Windows 中启动一次 Sticky Notes 应用。"
}

# 导出目录设置
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $defaultExportDir = Join-Path $scriptDir "StickyNotes_Export"
    $inputPath = Read-Host "请输入导出目录 (回车默认: .\StickyNotes_Export)"
    $OutputPath = if ([string]::IsNullOrWhiteSpace($inputPath)) { $defaultExportDir } else { $inputPath }
}

$targetDir = [System.IO.Path]::GetFullPath($OutputPath)
$attachmentsTargetDir = Join-Path $targetDir $MediaDirName
New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

# ----------------- 2. 文本解析与排版清洗 -----------------

function Convert-StickyTimestamp([object]$val) {
    if ($null -eq $val -or [string]::IsNullOrWhiteSpace([string]$val)) { return $null }
    $str = [string]$val
    if ($str -match '^\d+$') {
        $num = [double]$str
        if ($num -gt 1e14) {
            try { return ([DateTime]::SpecifyKind([DateTime]::MinValue.AddTicks([long]$num), [DateTimeKind]::Utc)).ToLocalTime() } catch {}
        }
        if ($num -gt 1e11) {
            try { return [DateTimeOffset]::FromUnixTimeMilliseconds([long]$num).ToLocalTime().DateTime } catch {}
        }
        if ($num -gt 1e9) {
            try { return [DateTimeOffset]::FromUnixTimeSeconds([long]$num).ToLocalTime().DateTime } catch {}
        }
    }
    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse($str, [ref]$parsed)) { return $parsed.ToLocalTime() }
    return $null
}

function Convert-StickyNotesContent([string]$source) {
    if ([string]::IsNullOrWhiteSpace($source)) { return "" }

    $text = $source

    # 1. 清除整行仅包含 \id=UUID 的空行占位符
    $text = [regex]::Replace($text, '(?m)^\s*\\?id=[0-9a-fA-F-]+\s*$', '')

    # 2. 清除行首的 \id=UUID 及其后最多 1 个分隔空格（保留用户输入的缩进空格）
    $text = $text -replace '\\?id=[0-9a-fA-F-]+ ?', ''

    # 3. 占位保护字面量双反斜杠
    $tokenSlash = [string][char]0xE000
    $text = $text.Replace('\\', $tokenSlash)

    # 4. 按行处理内联排版
    $lines = [regex]::Split($text, "\r?\n")
    $processedLines = [System.Collections.Generic.List[string]]::new()

    foreach ($rawLine in $lines) {
        $cur = $rawLine
        if ([string]::IsNullOrWhiteSpace($cur)) {
            $processedLines.Add("")
            continue
        }

        # 处理列表标记 \l / \l0（仅消耗最多 1 个分隔空格）
        $isList = $false
        if ($cur -match '\\l(?!\d)') {
            $isList = $true
            $cur = [regex]::Replace($cur, '\\l(?!\d) ?', '')
        }
        $cur = $cur -replace '\\l0 ?', ''

        # 处理成对标记（仅消耗标签后最多 1 个分隔空格）
        $cur = [regex]::Replace($cur, '\\b ?(.*?)\s*\\b0', '**$1**')
        $cur = [regex]::Replace($cur, '\\i ?(.*?)\s*\\i0', '*$1*')
        $cur = [regex]::Replace($cur, '\\strike ?(.*?)\s*\\strike0', '~~$1~~')
        $cur = [regex]::Replace($cur, '\\ul ?(.*?)\s*\\ul0', '<u>$1</u>')

        # 检查未显式闭合的标记
        $hasBold   = [bool]($cur -match '\\b(?!\d)')
        $hasItalic = [bool]($cur -match '\\i(?!\d)')
        $hasStrike = [bool]($cur -match '\\strike(?!\d)')
        $hasUl     = [bool]($cur -match '\\ul(?!\d)')

        # 清除残留指令（仅消耗最多 1 个分隔空格）
        $cur = [regex]::Replace($cur, '\\(b|i|ul|strike)(?!\d) ?', '')
        $cur = [regex]::Replace($cur, '\\(b|i|ul|strike)0 ?', '')

        # 仅清理行尾空格，完整保留前导缩进
        $cur = $cur.TrimEnd()

        # 若存在整行未闭合样式，分离前导缩进后再包裹正文，避免破坏缩进
        if ($hasBold -or $hasItalic -or $hasStrike -or $hasUl) {
            $indent = ""
            if ($cur -match '^([ \t]+)(.*)$') {
                $indent = $matches[1]
                $cur = $matches[2]
            }
            if ($hasBold)   { $cur = "**$cur**" }
            if ($hasItalic) { $cur = "*$cur*" }
            if ($hasStrike) { $cur = "~~$cur~~" }
            if ($hasUl)     { $cur = "<u>$cur</u>" }
            $cur = $indent + $cur
        }

        if ($isList) { $cur = "- $cur" }

        # 还原字面量反斜杠
        $cur = $cur.Replace($tokenSlash, '\')
        $processedLines.Add($cur)
    }

    $result = $processedLines -join "`n"
    $result = $result -replace '(\r?\n){3,}', "`n`n"
    return $result.TrimEnd()
}

# 提取纯净标题（严格先剥离 UUID，防止 \i 误杀 \id）
function Extract-PureTitle([string]$firstLine) {
    if ([string]::IsNullOrWhiteSpace($firstLine)) { return "无标题便签" }
    
    $clean = $firstLine
    # 第一步：必须绝对优先剥除 \id=UUID 标记
    $clean = $clean -replace '\\?id=[0-9a-fA-F-]+', ''
    
    # 第二步：剥除富文本控制符
    $clean = $clean -replace '\\(b|i|ul|strike|l)0?', ''
    $clean = $clean -replace '<[^>]+>', ''
    
    # 第三步：清理 Markdown 标记与 Windows 非法文件名字符
    $clean = $clean -replace '[\*~_#>`\[\]\(\)\\]', ''
    $clean = $clean -replace '[\\/:*?"<>|]', ' '
    $clean = $clean -replace '\s+', ' '
    $clean = $clean.Trim()

    if ($clean.Length -gt 40) {
        $clean = $clean.Substring(0, 40).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return "无标题便签"
    }
    return $clean
}

function Build-YamlFrontmatter([hashtable]$fields) {
    $yaml = @("---")
    foreach ($k in $fields.Keys) {
        $v = $fields[$k]
        if ($null -eq $v) { continue }
        if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) {
            $yaml += "$k`:"
            foreach ($item in $v) {
                $yaml += "  - `"$item`""
            }
        } elseif ($v -is [bool]) {
            $yaml += "$k`: $($v.ToString().ToLower())"
        } else {
            $escaped = [string]$v -replace '"', '\"'
            $yaml += "$k`: `"$escaped`""
        }
    }
    $yaml += "---"
    return ($yaml -join "`n")
}

# ----------------- 3. 读取数据库快照 -----------------

$origConsoleOut = [Console]::OutputEncoding
$origPsOut = $OutputEncoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$tempDir = Join-Path $env:TEMP ("StickyExportSnapshot_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    # 拷贝数据库及其 WAL/SHM 快照
    Get-ChildItem -Path $localStateDir -Filter "plum.sqlite*" | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $tempDir -Force
    }
    $tempPlum = Join-Path $tempDir "plum.sqlite"

    # 读取 Note 表数据
    $notesRaw = (& $sqliteExe -json $tempPlum "SELECT * FROM Note;") -join "`n"
    if ([string]::IsNullOrWhiteSpace($notesRaw) -or $notesRaw.Trim() -eq "[]") {
        Write-Host "便签数据库为空，无数据可导出。" -ForegroundColor Yellow
        return
    }
    $notes = @($notesRaw | ConvertFrom-Json)

    # 精准读取 Media 表（通过 ParentId 绑定便签 ID）
    $mediaMap = @{}
    try {
        $mediaSql = "SELECT ParentId, LocalFileRelativePath, AltText FROM Media WHERE ParentId IS NOT NULL;"
        $mediaRaw = (& $sqliteExe -json $tempPlum $mediaSql) -join "`n"
        if (-not [string]::IsNullOrWhiteSpace($mediaRaw)) {
            $mediaRecords = @($mediaRaw | ConvertFrom-Json)
            foreach ($m in $mediaRecords) {
                $noteParentId = [string]$m.ParentId
                if (-not $mediaMap.ContainsKey($noteParentId)) {
                    $mediaMap[$noteParentId] = [System.Collections.Generic.List[object]]::new()
                }
                $mediaMap[$noteParentId].Add($m)
            }
            Write-Host "[成功] 从 Media 表中检索到 $($mediaRecords.Count) 条附件索引记录" -ForegroundColor Green
        }
    } catch {
        Write-Warning "读取 Media 表失败: $_"
    }

    # ----------------- 4. 逐条生成 Markdown -----------------
    $count = 0
    $mediaExportCount = 0

    foreach ($note in $notes) {
        # 过滤已删除项
        if (-not $IncludeDeleted) {
            if ($note.PSObject.Properties['DeletedAt'] -and -not [string]::IsNullOrEmpty([string]$note.DeletedAt)) { continue }
            if ($note.PSObject.Properties['IsSoftDeleted'] -and $note.IsSoftDeleted -eq 1) { continue }
        }

        $noteId = [string]$note.Id
        $rawText = if ($note.PSObject.Properties['Text']) { [string]$note.Text } else { "" }

        # 解析时间戳
        $dtModified = Convert-StickyTimestamp $note.UpdatedAt
        $dtCreated  = Convert-StickyTimestamp $note.CreatedAt

        # 清洗并转换内联排版
        $body = Convert-StickyNotesContent $rawText

        # 提取文件名（杜绝 UUID 混入）
        $firstRawLine = ($rawText -split '\r?\n' | Where-Object { $_.Trim() -ne "" } | Select-Object -First 1)
        $title = Extract-PureTitle $firstRawLine

        # 查找当前笔记关联的媒体文件
        $imagesForNote = [System.Collections.Generic.List[string]]::new()
        if ($mediaMap.ContainsKey($noteId)) {
            New-Item -ItemType Directory -Path $attachmentsTargetDir -Force | Out-Null
            $idx = 1

            foreach ($mediaItem in $mediaMap[$noteId]) {
                $relPath = [string]$mediaItem.LocalFileRelativePath
                if ([string]::IsNullOrWhiteSpace($relPath)) { continue }

                # 清除首尾斜杠，防止路径被误判为根目录
                $cleanRelPath = $relPath.TrimStart('\', '/')
                $fileNameOnly = [System.IO.Path]::GetFileName($cleanRelPath)

                # 多路径冗余定位
                $candidatePaths = @(
                    (Join-Path $localStateDir $cleanRelPath),
                    (Join-Path $localStateDir "media\$cleanRelPath"),
                    (Join-Path $localStateDir "media\$fileNameOnly")
                )

                $sourceDiskFile = $null
                foreach ($cand in $candidatePaths) {
                    if (Test-Path $cand) {
                        $sourceDiskFile = $cand
                        break
                    }
                }

                if ($sourceDiskFile) {
                    $ext = [System.IO.Path]::GetExtension($sourceDiskFile)
                    if ([string]::IsNullOrWhiteSpace($ext)) { $ext = ".png" }
                    
                    $mediaFileName = "media_${noteId}_${idx}${ext}"
                    $destImgPath = Join-Path $attachmentsTargetDir $mediaFileName
                    
                    Copy-Item -Path $sourceDiskFile -Destination $destImgPath -Force
                    if ($dtModified) { [System.IO.File]::SetLastWriteTime($destImgPath, $dtModified) }
                    
                    $altText = if (-not [string]::IsNullOrWhiteSpace($mediaItem.AltText)) { $mediaItem.AltText } else { "$title 图片 $idx" }
                    $imagesForNote.Add("![$altText]($MediaDirName/$mediaFileName)")
                    $mediaExportCount++
                    # Write-Host "  -> [导出附件成功] 便签: '$title' -> $mediaFileName" -ForegroundColor Cyan
                    $idx++
                } else {
                    Write-Warning "未在磁盘找到该附件: $cleanRelPath (已搜索: $localStateDir\media)"
                }
            }
        }

        # 如果既无文字又无图片则跳过
        if ([string]::IsNullOrWhiteSpace($body) -and $imagesForNote.Count -eq 0) { continue }

        # 图片悬浮置顶（Frontmatter 与正文之间）
        $topMediaBlock = if ($imagesForNote.Count -gt 0) {
            ($imagesForNote -join "`n`n") + "`n`n"
        } else { "" }

        $theme = if ($note.PSObject.Properties['Theme']) { [string]$note.Theme } else { "Default" }

        # 构造 YAML Frontmatter
        $frontmatterFields = [ordered]@{
            "sticky_id"   = $noteId
            "created"     = if ($dtCreated) { $dtCreated.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            "title"       = $title
            "color"       = $theme
            "updated"     = if ($dtModified) { $dtModified.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            "tags"        = @("sticky-notes", "color/$($theme.ToLower())")
        }
        $frontmatterText = Build-YamlFrontmatter $frontmatterFields

        # 完整结构：Frontmatter -> 置顶图片 -> 正文
        $finalMarkdown = $frontmatterText + "`n`n" + $topMediaBlock + $body

        # 文件命名与防重
        $fileName = "$title.md"
        $destPath = Join-Path $targetDir $fileName
        $seq = 2
        while (Test-Path $destPath) {
            $destPath = Join-Path $targetDir ("$title ($seq).md")
            $seq++
        }

        [System.IO.File]::WriteAllText($destPath, $finalMarkdown + "`n", (New-Object System.Text.UTF8Encoding($false)))

        # 同步物理时间
        if ($dtCreated)  { [System.IO.File]::SetCreationTime($destPath, $dtCreated) }
        if ($dtModified) { [System.IO.File]::SetLastWriteTime($destPath, $dtModified) }
        $count++
    }

    Write-Host "`n[导出完成]" -ForegroundColor Green
    Write-Host "成功导出便签: $count 篇" -ForegroundColor Cyan
    Write-Host "成功导出图片: $mediaExportCount 个" -ForegroundColor Cyan
    Write-Host "存放目录: $targetDir" -ForegroundColor Gray
}
finally {
    [Console]::OutputEncoding = $origConsoleOut
    $OutputEncoding = $origPsOut

    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}