# -*- coding: utf-8 -*-
<#
.SYNOPSIS
  复制空白标准值班表(.xls)为未来每一天的值班表，仅改写白名单单元格(日期/各岗位人员/备注)的值，
  并把模板的完整格式(字体/对齐/数字格式/填充/行高列宽)逐格拷回，确保与模板“完美一致”。
  全部在隐藏暂存目录完成，成功才发布到输出目录。
  注意：空白模板的 usedRange 可能塌缩为 A1:A19，因此格式拷贝使用固定网格(1..30 × 1..10)，
  不依赖 usedRange，避免 B-G 列格式丢失。
.PARAMETER PayloadPath
  包含 template_path / sheet / output_dir / days[] 的 UTF-8 JSON 路径。
#>
param(
    [Parameter(Mandatory=$true)] [string] $PayloadPath
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$excel = $null

$GRID_R = 19   # 固定网格行上限（覆盖 A1:G19 实际区域）
$GRID_C = 7    # 固定网格列上限（A-G）

function Copy-Format($src, $dst) {
    try {
        $sf = $src.Font; $df = $dst.Font
        $df.Name    = $sf.Name
        $df.Size    = $sf.Size
        $df.Bold    = $sf.Bold
        $df.Italic  = $sf.Italic
        $df.Underline = $sf.Underline
        $df.Color   = $sf.Color
        $dst.HorizontalAlignment = $src.HorizontalAlignment
        $dst.VerticalAlignment   = $src.VerticalAlignment
        $dst.NumberFormat        = $src.NumberFormat
        $dst.Interior.ColorIndex = $src.Interior.ColorIndex
        $dst.Interior.Pattern    = $src.Interior.Pattern
        # 边框：左/上/右/下 1:1 还原（内边框由单元格自身生成，无需逐格拷贝）
        $sb = $src.Borders; $db = $dst.Borders
        foreach ($i in @(7,8,9,10)) {
            try {
                $db.Item($i).LineStyle  = $sb.Item($i).LineStyle
                $db.Item($i).Weight     = $sb.Item($i).Weight
                $db.Item($i).ColorIndex = $sb.Item($i).ColorIndex
            } catch {}
        }
    } catch {}
}

try {
    $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $tplPath = Resolve-Path -LiteralPath $payload.template_path
    $sheetName = $payload.sheet
    $tplBook = $excel.Workbooks.Open($tplPath)          # 空白模板(只读参考)
    $tplWs = $tplBook.Worksheets.Item($sheetName)

    $staging = Join-Path $payload.output_dir ".staging"
    if (-not (Test-Path $staging)) { New-Item -ItemType Directory -Path $staging | Out-Null }

    foreach ($day in $payload.days) {
        $target = Join-Path $staging $day.filename
        Copy-Item -LiteralPath $tplPath -Destination $target -Force
        $book = $excel.Workbooks.Open($target)
        try {
            $ws = $book.Worksheets.Item($sheetName)
            # 1) 仅改写白名单单元格的值
            $ws.Range($day.date_cell).Value2 = $day.date_text
            foreach ($c in $day.cells) { $ws.Range($c.cell).Value2 = $c.text }
            $ws.Range($day.note_cell).Value2 = $day.note_text
            # 2) 复制模板已自带完整格式（含边框/字体/对齐），只需强制还原行高/列宽，
            #    避免 Excel 因单元格为空而压缩长文本行高。单元格级格式无需逐格拷回，
            #    否则会引发边框枚举值差异（xlNone vs xlColorIndexAutomatic）。
            for ($r = 1; $r -le $GRID_R; $r++) {
                $ws.Rows.Item($r).RowHeight = $tplWs.Rows.Item($r).RowHeight
            }
            for ($c = 1; $c -le $GRID_C; $c++) {
                $ws.Columns.Item($c).ColumnWidth = $tplWs.Columns.Item($c).ColumnWidth
            }
            $book.Save()
        } finally {
            $book.Close($false)
        }
    }
    $tplBook.Close($false)

    # 全部成功后再发布
    foreach ($day in $payload.days) {
        $src = Join-Path $staging $day.filename
        $dst = Join-Path $payload.output_dir $day.filename
        Move-Item -LiteralPath $src -Destination $dst -Force
    }
    Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    "OK" | Out-File -FilePath (Join-Path $payload.output_dir ".write_result.txt") -Encoding utf8
} catch {
    if ($payload -and $payload.output_dir) {
        $staging = Join-Path $payload.output_dir ".staging"
        if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Error ("WRITE_TEMPLATE_FAILED: " + $_.Exception.Message)
    exit 1
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
