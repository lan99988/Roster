# -*- coding: utf-8 -*-
<#
.SYNOPSIS
  严格比对生成值与班表与空白标准模板的格式指纹：工作表结构、合并区域、固定网格内每个单元格的格式，
  以及非白名单单元格的值。任何非白名单差异都视为失败（保证完美一致）。
  空白模板的 usedRange 可能塌缩为 A1:A19，因此统一用固定网格(1..30 × 1..10)读写，不依赖 usedRange。
.PARAMETER TemplatePath
.PARAMETER OutputPaths  (逗号分隔)
.PARAMETER AllowedCells (逗号分隔，如 A3,A19,B5,D5)
.PARAMETER ReportPath
#>
param(
    [Parameter(Mandatory=$true)] [string] $TemplatePath,
    [Parameter(Mandatory=$true)] [string] $OutputPaths,
    [Parameter(Mandatory=$true)] [string] $AllowedCells,
    [Parameter(Mandatory=$true)] [string] $ReportPath
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$excel = $null
$GRID_R = 30
$GRID_C = 10

function SafeInt($v) {
    if ($v -is [System.DBNull] -or $v -eq $null) { return 0 }
    try { return [int]$v } catch { return 0 }
}
function SafeDouble($v) {
    if ($v -is [System.DBNull] -or $v -eq $null) { return 0.0 }
    try { return [double]$v } catch { return 0.0 }
}
function SafeStr($v) {
    if ($v -is [System.DBNull] -or $v -eq $null) { return "" }
    return [string]$v
}
function BorderSig($b) {
    return "{0},{1},{2}" -f (SafeInt $b.LineStyle), (SafeInt $b.Weight), (SafeInt $b.ColorIndex)
}

function Get-Fingerprint($path, $sheetName) {
    $book = $excel.Workbooks.Open($path)
    try {
        $ws = $book.Worksheets.Item($sheetName)
        $merged = @()
        foreach ($m in $ws.UsedRange.MergeAreas) {
            if ($m -and $m.Address) { $merged += $m.Address() }
        }
        $cells = @{}
        $rowH = @{}
        $colW = @{}
        for ($r = 1; $r -le $GRID_R; $r++) {
            $rowH[$r] = SafeDouble $ws.Rows.Item($r).RowHeight
            for ($c = 1; $c -le $GRID_C; $c++) {
                $cell = $ws.Cells.Item($r, $c)
                $addr = $cell.Address($false, $false)
                $f = $cell.Font
                $b = $cell.Borders
                $cells[$addr] = [ordered]@{
                    value = SafeStr $cell.Value2
                    bold = [bool]$f.Bold
                    size = SafeDouble $f.Size
                    color = SafeInt $f.Color
                    fill = SafeInt $cell.Interior.ColorIndex
                    numfmt = SafeStr $cell.NumberFormat
                    halign = SafeInt $cell.HorizontalAlignment
                    borderL = BorderSig $b.Item(7)
                    borderT = BorderSig $b.Item(8)
                    borderR = BorderSig $b.Item(9)
                    borderB = BorderSig $b.Item(10)
                    borderV = BorderSig $b.Item(11)
                    borderH = BorderSig $b.Item(12)
                }
            }
        }
        for ($c = 1; $c -le $GRID_C; $c++) {
            $colW[$c] = SafeDouble $ws.Columns.Item($c).ColumnWidth
        }
        return [ordered]@{
            sheetNames = @($book.Worksheets | ForEach-Object { $_.Name })
            merged = $merged
            cells = $cells
            rowH = $rowH
            colW = $colW
        }
    } finally {
        $book.Close($false)
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
    $sheetName = "8月排班表"
    $allowedSet = @{}
    foreach ($a in ($AllowedCells -split ',')) { $allowedSet[$a.Trim()] = $true }
    $tpl = Get-Fingerprint (Resolve-Path $TemplatePath) $sheetName
    $diffs = @()
    foreach ($op in ($OutputPaths -split ',')) {
        $op = $op.Trim()
        if (-not $op) { continue }
        $out = Get-Fingerprint (Resolve-Path $op) $sheetName
        $leaf = Split-Path $op -Leaf
        if (($out.sheetNames -join '|') -ne ($tpl.sheetNames -join '|')) {
            $diffs += "${leaf}: 工作表结构不一致 $($out.sheetNames -join ',')"
        }
        if (($out.merged -join '|') -ne ($tpl.merged -join '|')) {
            $diffs += "${leaf}: 合并区域不一致"
        }
        # 行高 / 列宽：全部 1:1 还原，含白名单格的行/列格式也应一致
        foreach ($r in $tpl.rowH.Keys) {
            if ($out.rowH.ContainsKey($r) -and $out.rowH[$r] -ne $tpl.rowH[$r]) {
                $diffs += "${leaf}: 第${r}行行高不一致 ($($out.rowH[$r]) vs $($tpl.rowH[$r]))"
            }
        }
        foreach ($c in $tpl.colW.Keys) {
            if ($out.colW.ContainsKey($c) -and $out.colW[$c] -ne $tpl.colW[$c]) {
                $diffs += "${leaf}: 第${c}列列宽不一致 ($($out.colW[$c]) vs $($tpl.colW[$c]))"
            }
        }
        # 逐格：非白名单格必须值与格式一致；白名单格只跳过「值」比对，格式仍必须一致
        foreach ($addr in $tpl.cells.Keys) {
            $tc = $tpl.cells[$addr]; $oc = $out.cells[$addr]
            if (-not $oc) { $diffs += "${leaf}: 缺单元格 $addr"; continue }
            $isAllowed = $allowedSet.ContainsKey($addr)
            if (-not $isAllowed -and $oc.value -ne $tc.value) {
                $diffs += "${leaf}: 单元格 $addr 值被非授权修改 '$($oc.value)'（应为 '$($tc.value)'）"
            }
            foreach ($k in @('bold','size','color','fill','numfmt','halign','borderL','borderT','borderR','borderB')) {
                if ($oc[$k] -ne $tc[$k]) {
                    $diffs += ("${leaf}: 单元格 {0} 格式[{1}]不一致 ({2} vs {3})" -f $addr, $k, $oc[$k], $tc[$k])
                }
            }
        }
        # 输出中多出的单元格（模板网格外）若含值，也视为异常
        foreach ($addr in $out.cells.Keys) {
            if ($tpl.cells.ContainsKey($addr)) { continue }
            $oc = $out.cells[$addr]
            if ($oc.value -ne "" -and $oc.value -ne $null) {
                $diffs += "${leaf}: 单元格 $addr 模板外出现非空格 '$($oc.value)'"
            }
        }
    }
    $report = [ordered]@{ verified = ($diffs.Count -eq 0); diffs = $diffs }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
} catch {
    Write-Error ("FINGERPRINT_FAILED: " + $_.Exception.Message)
    exit 1
} finally {
    if ($excel) { try { $excel.Quit() } catch {} [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
