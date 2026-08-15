$ErrorActionPreference = 'Stop'
# build_blank_template.ps1 <sourceXls> <outXls>
# 从最新历史值班表生成「无名字」的标准空白模板：
#   - 保留标题行(R3)的 B/C/D/G 时间段文本与全部格式
#   - 仅清除岗位数据行(R4..R18)的 B/C/D/G 指派列
#   - 清除日期单元格 A3 的值（保留格式）
#   - 归一化备注(A19)，删除每日变化的条目 5、6
#   - 完整保留原表行高/列宽/边框/字体等格式
$src = $args[0]
$out = $args[1]
$excel = $null
function ClearCell($ws, $r, $c) {
    $cell = $ws.Cells.Item($r, $c)
    try {
        $ma = $cell.MergeArea
        if ($ma.Count -gt 1) {
            if ($ma.Row -ne $r -or $ma.Column -ne $c) { return }  # not top-left, skip
        }
    } catch {}
    $cell.Value2 = ""   # empty string keeps the cell "used" so borders/format survive SaveAs
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($src)
    $ws = $wb.Worksheets.Item(1)
    $ur = $ws.UsedRange
    $r1 = $ur.Row
    $r2 = $ur.Row + $ur.Rows.Count - 1
    $c1 = $ur.Column
    $c2 = $ur.Column + $ur.Columns.Count - 1

    # 先记录原表行高，清除后强制写回，避免 Excel 因单元格为空而压缩长文本行
    $origRowHeights = @{}
    for ($r = 1; $r -le $r2; $r++) { $origRowHeights[$r] = $ws.Rows.Item($r).RowHeight }

    # 仅清除岗位数据行 R4..R18 的指派列 B/C/D/G；保留 R3 标题行
    foreach ($c in @(2,3,4,7)) {       # B,C,D,G
        if ($c -ge $c1 -and $c -le $c2) {
            for ($r = [Math]::Max(4, $r1); $r -le [Math]::Min(18, $r2); $r++) { ClearCell $ws $r $c }
        }
    }
    # 清除日期单元格 A3 的值（格式保留，由写入时填新日期）
    if ($r1 -le 3 -and $r2 -ge 3) { ClearCell $ws 3 1 }

    # 归一化备注(A19)：删掉动态项 5、6（调休/请假等每日变化条目），只留常设 1-4、7-9
    $noteCell = $ws.Cells.Item(19, 1)
    $note = [string]$noteCell.Text
    if ($note) {
        $note = $note -replace '(?<=；)\s*[56][^；]*；', ''
        $noteCell.Value2 = $note
    }

    # 强制还原行高
    for ($r = 1; $r -le $r2; $r++) { $ws.Rows.Item($r).RowHeight = $origRowHeights[$r] }

    $dir = Split-Path $out
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $wb.SaveAs($out, 56)   # xlExcel8
    $wb.Close($false)
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Host "BLANK_TEMPLATE_BUILT $out"
} catch {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Host "ERR: $_"
    exit 1
} finally {
    if ($excel) { $excel.Quit() }
}
