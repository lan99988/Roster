# Template-Preserving Roster Output Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate three future `.xls` duty rosters by copying the latest complete real roster template with Microsoft Excel and changing only approved date, personnel, and same-day note cells.

**Architecture:** Keep parsing and scheduling deterministic in Python, but replace the generated daily workbook path with a template-aware demand model plus a PowerShell Excel COM writer. The writer works in a hidden staging directory, verifies a strict format fingerprint against the source template, and publishes all three files only after every file passes business, structural, and visual-export checks.

**Tech Stack:** Python 3.10+, `xlrd`, `unittest`, PowerShell 5.1, Microsoft Excel COM, existing `@oai/artifact-tool` report writer.

---

### Task 1: Model and Select a Complete Template

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/scripts/template_model.py`
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_template_model.py`

- [ ] **Step 1: Write failing tests for completeness and selection**

```python
def test_selects_latest_complete_template():
    candidates = [
        {"path": "8.11.xls", "date": "2026-08-11", "labels": ["游客服务驿站"], "print_area": "$A$1:$H$2"},
        {"path": "8.13.xls", "date": "2026-08-13", "labels": REQUIRED_TEMPLATE_LABELS, "print_area": ""},
    ]
    selected, warnings = select_template(candidates)
    assert selected["path"] == "8.13.xls"
    assert "8.11.xls" in warnings[0]


def test_rejects_template_missing_required_label():
    with self.assertRaisesRegex(TemplateError, "缺少岗位行"):
        validate_template({"labels": ["游客服务驿站"], "sheet_name": "8月排班表"})
```

- [ ] **Step 2: Run the new tests and verify RED**

Run:

```powershell
python -m unittest discover -s .\.workbuddy\skills\duty-roster-generator\tests -p "test_template_model.py" -v
```

Expected: FAIL because `template_model` does not exist.

- [ ] **Step 3: Implement the pure template model**

Define:

```python
REQUIRED_TEMPLATE_LABELS = (
    "游客服务驿站",
    "门口指引、闸机",
    "正堂音响/园区广播",
    "演艺控场",
    "视频号宣传",
    "开衙/巡府",
    "免费讲解",
    "侨批展邮政文创摊位",
    "行政事务/巡园/机动岗位",
)


class TemplateError(ValueError):
    pass


def is_complete_template(metadata):
    labels = set(metadata.get("labels", ()))
    return metadata.get("sheet_name") == "8月排班表" and set(REQUIRED_TEMPLATE_LABELS) <= labels


def select_template(candidates):
    complete = [item for item in candidates if is_complete_template(item)]
    if not complete:
        raise TemplateError("过往三天中没有结构完整的值班表模板")
    selected = max(complete, key=lambda item: item["date"])
    warnings = [f'{item["path"]} 结构不完整，未作为模板' for item in candidates if item not in complete]
    return selected, warnings
```

- [ ] **Step 4: Run tests and verify GREEN**

Expected: all `test_template_model.py` tests pass.

- [ ] **Step 5: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/template_model.py .workbuddy/skills/duty-roster-generator/tests/test_template_model.py
git commit -m "feat: select latest complete roster template"
```

### Task 2: Inspect Real `.xls` Structure Through Excel

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/scripts/inspect_excel_template.ps1`
- Create: `.workbuddy/skills/duty-roster-generator/scripts/excel_bridge.py`
- Modify: `.workbuddy/skills/duty-roster-generator/tests/test_template_model.py`

- [ ] **Step 1: Add a failing bridge contract test**

```python
def test_inspector_returns_sheet_labels_and_print_area(tmp_path):
    result = inspect_template(FIXTURE_XLS, command_runner=fake_excel_result)
    assert result["sheet_name"] == "8月排班表"
    assert result["used_range"] == "$A$3:$G$19"
    assert "游客服务驿站" in result["labels"]
```

- [ ] **Step 2: Verify RED**

Expected: FAIL because `inspect_template` is missing.

- [ ] **Step 3: Implement `inspect_excel_template.ps1`**

The script accepts `-WorkbookPath` and `-OutputJson`, opens Excel read-only, selects `8月排班表`, and writes UTF-8 JSON containing:

```powershell
$result = [ordered]@{
    sheet_name = $sheet.Name
    used_range = $sheet.UsedRange.Address()
    print_area = $sheet.PageSetup.PrintArea
    labels = @($sheet.Range('A1:A80').Value2 | ForEach-Object { if ($_ -is [array]) { $_[0] } else { $_ } })
    sheet_names = @($book.Worksheets | ForEach-Object { $_.Name })
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
```

Always close the workbook and call `Excel.Application.Quit()` in `finally`.

- [ ] **Step 4: Implement the Python bridge**

```python
def inspect_template(path, runner=subprocess.run):
    with tempfile.TemporaryDirectory() as temporary:
        output = Path(temporary) / "template.json"
        completed = runner([
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(SCRIPTS / "inspect_excel_template.ps1"),
            "-WorkbookPath", str(Path(path).resolve()),
            "-OutputJson", str(output),
        ], capture_output=True, text=True, encoding="utf-8", timeout=60)
        if completed.returncode:
            raise TemplateError(completed.stderr.strip() or "Excel 模板检查失败")
        return json.loads(output.read_text(encoding="utf-8-sig"))
```

- [ ] **Step 5: Run unit tests and a real inspection smoke test**

Run:

```powershell
python -m unittest discover -s .\.workbuddy\skills\duty-roster-generator\tests -p "test_template_model.py" -v
powershell -NoProfile -ExecutionPolicy Bypass -File .\.workbuddy\skills\duty-roster-generator\scripts\inspect_excel_template.ps1 -WorkbookPath .\输入\过往三天\8.13值班表.xls -OutputJson $env:TEMP\roster-template.json
```

Expected: tests pass and JSON reports `A3:G19` plus all required labels.

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/inspect_excel_template.ps1 .workbuddy/skills/duty-roster-generator/scripts/excel_bridge.py .workbuddy/skills/duty-roster-generator/tests/test_template_model.py
git commit -m "feat: inspect roster templates with Excel"
```

### Task 3: Replace Abstract Posts With Template Cell Demands

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/config/template_map.json`
- Create: `.workbuddy/skills/duty-roster-generator/scripts/template_demands.py`
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_template_demands.py`
- Modify: `.workbuddy/skills/duty-roster-generator/scripts/three_day_scheduler.py`

- [ ] **Step 1: Write failing mapping tests**

```python
def test_maps_template_roles_to_cells():
    demands = build_template_demands(TEMPLATE_MAP)
    assert demands["游客服务驿站:上午"]["cell"] == "B4"
    assert demands["游客服务驿站:下午"]["cell"] == "D4"
    assert demands["门口指引、闸机:全天机动"]["cell"] == "G5"
    assert demands["行政事务/巡园/机动:下午"]["cell"] == "D18"


def test_every_required_template_cell_receives_names():
    result = schedule_three_days(
        dates=[date(2026, 8, 14)],
        staff=[
            {"name": "女甲", "tags": ["女", "主持人"]},
            {"name": "女乙", "tags": ["女", "主持人"]},
            {"name": "女丙", "tags": ["女", "宣传"]},
            {"name": "女丁", "tags": ["女"]},
            {"name": "技师甲", "tags": ["技术"]},
            {"name": "技师乙", "tags": ["技术"]},
            {"name": "负责人甲", "tags": ["负责人"]},
            {"name": "负责人乙", "tags": ["负责人"]},
            *[{"name": f"通用{i}", "tags": []} for i in range(1, 16)],
        ],
        posts=build_template_demands(TEMPLATE_MAP),
        changes={"unavailable": [], "assignments": [], "warnings": []},
        history={},
        rest_target=2,
    )
    assert not [item for item in result["days"][0]["cells"] if item["required"] and not item["names"]]
```

- [ ] **Step 2: Verify RED**

Expected: FAIL because cell demands are not produced.

- [ ] **Step 3: Add explicit template mapping**

`template_map.json` must define the real current anchors:

```json
{
  "sheet": "8月排班表",
  "date_cell": "A3",
  "note_cell": "A19",
  "cells": [
    {"key":"游客服务驿站:上午","label":"游客服务驿站","cell":"B4","count":2,"required_tags":["女"]},
    {"key":"游客服务驿站:下午","label":"游客服务驿站","cell":"D4","count":2,"required_tags":["女"]},
    {"key":"门口指引、闸机:开馆","label":"门口指引、闸机","cell":"B5","count":1},
    {"key":"门口指引、闸机:关馆及晚班","label":"门口指引、闸机","cell":"D5","count":4},
    {"key":"门口指引、闸机:全天机动","label":"门口指引、闸机","cell":"G5","count":6},
    {"key":"正堂音响/园区广播:上午","label":"正堂音响/园区广播","cell":"B7","count":1,"required_tags":["技术"]},
    {"key":"正堂音响/园区广播:下午","label":"正堂音响/园区广播","cell":"D7","count":1,"required_tags":["技术"]},
    {"key":"演艺控场:上午","label":"演艺控场","cell":"C8","count":1,"required_tags":["主持人"]},
    {"key":"演艺控场:下午","label":"演艺控场","cell":"D8","count":1,"required_tags":["主持人"]},
    {"key":"视频号宣传:白班","label":"视频号宣传","cell":"C9","count":1,"required_tags":["宣传"]},
    {"key":"开衙/巡府","label":"开衙/巡府","cell":"C10","count":2,"required_tags":["主持人"]},
    {"key":"免费讲解:11:30","label":"免费讲解 11:30 16:30","cell":"C11","count":1,"required_tags":["主持人"]},
    {"key":"免费讲解:16:30","label":"免费讲解 11:30 16:30","cell":"D11","count":1,"required_tags":["主持人"]},
    {"key":"侨批展邮政文创摊位:上午","label":"侨批展邮政文创摊位","cell":"B14","count":1,"required_tags":["女"]},
    {"key":"侨批展邮政文创摊位:协同","label":"侨批展邮政文创摊位","cell":"C14","count":1,"required_tags":["女"]},
    {"key":"侨批展邮政文创摊位:下午","label":"侨批展邮政文创摊位","cell":"D14","count":1,"required_tags":["女"]},
    {"key":"行政事务/巡园/机动:上午","label":"行政事务/巡园/机动岗位","cell":"B18","count":1,"required_tags":["负责人"]},
    {"key":"行政事务/巡园/机动:下午","label":"行政事务/巡园/机动岗位","cell":"D18","count":1,"required_tags":["负责人"]}
  ]
}
```

- [ ] **Step 4: Extend the scheduler to solve cell demands**

Return each day as:

```python
{
    "date": "2026-08-14",
    "cells": [{"key": demand["key"], "cell": demand["cell"], "names": selected_names, "required": True}],
    "rest": [{"name": "通用14", "kind": "调休"}],
    "issues": [],
}
```

Use the existing backtracking solver, but make every template cell slot a unit. A person may be reused only when the mapped time windows do not overlap and the template explicitly models the transition.

- [ ] **Step 5: Run scheduler and mapping tests**

Expected: mapping tests and all existing scheduler tests pass.

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/config/template_map.json .workbuddy/skills/duty-roster-generator/scripts/template_demands.py .workbuddy/skills/duty-roster-generator/scripts/three_day_scheduler.py .workbuddy/skills/duty-roster-generator/tests/test_template_demands.py
git commit -m "feat: schedule against real template cells"
```

### Task 4: Build Cell Text Without Changing Cell Formatting

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/scripts/template_text.py`
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_template_text.py`

- [ ] **Step 1: Write failing text-rendering tests**

```python
def test_preserves_role_suffix_while_replacing_names():
    assert render_cell("惠君、洁娜（开六房、机动）", ["员工甲", "员工乙"], "游客服务驿站:上午") == "员工甲、员工乙（开六房、机动）"


def test_builds_date_header_with_weekday():
    assert render_date(date(2026, 8, 14)) == "8月14日周五"


def test_note_replaces_only_daily_leave_clause():
    text = render_note(BASE_NOTE, [{"name":"员工甲","kind":"请假（下午）"}])
    assert "员工甲下午请假" in text
    assert "文创摊位人员兼顾后闸机" in text
```

- [ ] **Step 2: Verify RED**

- [ ] **Step 3: Implement explicit renderers per mapped cell key**

Use a renderer table rather than free-form string replacement:

```python
RENDERERS = {
    "游客服务驿站:上午": lambda names: f'{"、".join(names)}（开六房、机动）',
    "游客服务驿站:下午": lambda names: f'{"、".join(names)}（机动）',
    "正堂音响/园区广播:上午": lambda names: f'{names[0]}（开六房）',
    "正堂音响/园区广播:下午": lambda names: f'{names[0]}（关六房）',
    "免费讲解:11:30": lambda names: f'11:30 {names[0]}',
    "免费讲解:16:30": lambda names: f'16:30 {names[0]}',
}
```

Keys with complex multi-line duties (`D5`) get dedicated functions with exact line breaks.

- [ ] **Step 4: Run tests and verify GREEN**

- [ ] **Step 5: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/template_text.py .workbuddy/skills/duty-roster-generator/tests/test_template_text.py
git commit -m "feat: render template cell text deterministically"
```

### Task 5: Copy and Edit Workbooks With Excel COM

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/scripts/write_template_roster.ps1`
- Modify: `.workbuddy/skills/duty-roster-generator/scripts/excel_bridge.py`
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_excel_bridge.py`

- [ ] **Step 1: Write a failing command contract test**

```python
def test_writer_uses_hidden_staging_directory(tmp_path):
    write_template_rosters(TEMPLATE, PAYLOAD, tmp_path, runner=fake_runner)
    assert fake_runner.args[-2:] == ["-PayloadPath", str(payload_path)]
    assert not list(tmp_path.glob("*.xls"))
```

- [ ] **Step 2: Verify RED**

- [ ] **Step 3: Implement the PowerShell writer**

For each day:

```powershell
$target = Join-Path $StagingDirectory ($day.date + '值班表.xls')
Copy-Item -LiteralPath $TemplatePath -Destination $target
$book = $excel.Workbooks.Open($target)
$sheet = $book.Worksheets.Item($payload.template.sheet)
$sheet.Range($payload.template.date_cell).Value2 = $day.date_text
foreach ($cell in $day.cells) {
    $sheet.Range($cell.cell).Value2 = $cell.text
}
$sheet.Range($payload.template.note_cell).Value2 = $day.note_text
$book.Save()
```

Set `$excel.Visible = $false`, `$excel.DisplayAlerts = $false`, and release every COM object in `finally`. Never assign `.Font`, `.Interior`, `.Borders`, `.RowHeight`, `.ColumnWidth`, or `.PageSetup`.

- [ ] **Step 4: Implement Python orchestration and error translation**

The Python bridge writes one UTF-8 payload JSON, invokes PowerShell with argument arrays, and converts non-zero exit status into `TemplateError` without exposing a traceback to daily users.

- [ ] **Step 5: Run unit tests and a one-file real smoke test in a temporary directory**

Expected: the copied `.xls` opens successfully in Excel and the template source remains byte-for-byte unchanged.

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/write_template_roster.ps1 .workbuddy/skills/duty-roster-generator/scripts/excel_bridge.py .workbuddy/skills/duty-roster-generator/tests/test_excel_bridge.py
git commit -m "feat: write roster copies with Excel"
```

### Task 6: Verify a Strict Format Fingerprint

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/scripts/fingerprint_excel.ps1`
- Create: `.workbuddy/skills/duty-roster-generator/scripts/format_fingerprint.py`
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_format_fingerprint.py`

- [ ] **Step 1: Write failing fingerprint comparison tests**

```python
def test_allows_only_whitelisted_value_changes():
    compare_fingerprints(TEMPLATE_FP, OUTPUT_FP, allowed_cells={"A3", "B4", "D4", "A19"})


def test_rejects_column_width_change():
    changed = deepcopy(TEMPLATE_FP)
    changed["columns"]["B"]["width"] += 1
    with self.assertRaisesRegex(FormatMismatch, "列宽"):
        compare_fingerprints(TEMPLATE_FP, changed, allowed_cells={"A3"})
```

- [ ] **Step 2: Verify RED**

- [ ] **Step 3: Export the Excel fingerprint**

The PowerShell script records sheet order, used range, merge areas, row heights, column widths, hidden states, page setup, formulas, values, and per-cell formatting fields. Serialize colors as numeric Excel color values and borders by edge/style/weight/color.

- [ ] **Step 4: Implement comparison with a cell whitelist**

```python
class FormatMismatch(ValueError):
    pass


def compare_fingerprints(template, output, allowed_cells):
    for section in ("sheet_names", "merged", "rows", "columns", "page_setup"):
        if template[section] != output[section]:
            raise FormatMismatch(f"模板格式不一致：{section}")
    for address, expected in template["cells"].items():
        actual = output["cells"][address]
        if expected["format"] != actual["format"]:
            raise FormatMismatch(f"单元格 {address} 格式不一致")
        if address not in allowed_cells and expected["value"] != actual["value"]:
            raise FormatMismatch(f"单元格 {address} 出现非授权内容变化")
```

- [ ] **Step 5: Run unit and real Excel fingerprint tests**

Expected: editing `B4.Value2` passes; changing `B4.Font.Bold` fails.

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/fingerprint_excel.ps1 .workbuddy/skills/duty-roster-generator/scripts/format_fingerprint.py .workbuddy/skills/duty-roster-generator/tests/test_format_fingerprint.py
git commit -m "feat: verify roster format fingerprints"
```

### Task 7: Make Three-Day Publication Transactional

**Files:**
- Modify: `.workbuddy/skills/duty-roster-generator/scripts/roster_pipeline.py`
- Modify: `.workbuddy/skills/duty-roster-generator/tests/test_pipeline.py`

- [ ] **Step 1: Write failing transaction tests**

```python
def test_partial_writer_failure_publishes_no_daily_files(tmp_path):
    writer = FakeWriter(fail_on_day=2)
    result = run_pipeline(tmp_path, "", now=NOW, template_writer=writer)
    assert result["status"] == "blocked"
    assert not list(Path(result["output_dir"]).glob("*值班表.xls"))


def test_success_publishes_exactly_three_daily_files(tmp_path):
    result = run_pipeline(tmp_path, "", now=NOW, template_writer=FakeWriter())
    assert len(list(Path(result["output_dir"]).glob("????-??-??值班表.xls"))) == 3
```

- [ ] **Step 2: Verify RED**

- [ ] **Step 3: Integrate template selection, cell scheduling, writing and fingerprint checks**

Create `batch/.staging`, write and verify all files there, then publish with `Path.replace()` only after all three pass. On any error, delete staging and generate only `排班说明与告警.xlsx` plus `运行摘要.json`.

- [ ] **Step 4: Add summary fields**

`运行摘要.json` must include:

```json
{
  "template": {"path":"输入/过往三天/8.13值班表.xls","date":"2026-08-13","fingerprint":"sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"},
  "daily_files": ["2026-08-14值班表.xls","2026-08-15值班表.xls","2026-08-16值班表.xls"],
  "format_verification": "passed"
}
```

- [ ] **Step 5: Run pipeline tests**

Expected: partial failure leaves zero daily files; success publishes exactly three.

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/scripts/roster_pipeline.py .workbuddy/skills/duty-roster-generator/tests/test_pipeline.py
git commit -m "feat: publish template rosters transactionally"
```

### Task 8: Add Real Excel Integration and Visual Verification

**Files:**
- Create: `.workbuddy/skills/duty-roster-generator/tests/test_excel_integration.py`
- Create: `.workbuddy/skills/duty-roster-generator/scripts/export_roster_preview.ps1`

- [ ] **Step 1: Add a real Excel integration test guarded by Excel availability**

The test copies `8.13值班表.xls` into a temporary fixture directory, writes one date and representative personnel cells, compares fingerprints, reopens the output, and asserts the values. Skip only when `New-Object -ComObject Excel.Application` is unavailable.

- [ ] **Step 2: Run the test and verify it fails before the writer exists**

- [ ] **Step 3: Implement preview export**

Open each output with Excel, export only `8月排班表` to PDF using the verified template print area, then render the first page to PNG with bundled Poppler. Preview artifacts remain under the batch `.verification` directory and are not included as normal user outputs.

- [ ] **Step 4: Run the full suite**

```powershell
python -m unittest discover -s .\.workbuddy\skills\duty-roster-generator\tests -p "test_*.py" -v
```

Expected: all unit tests pass and real Excel integration passes on this machine.

- [ ] **Step 5: Generate a real three-day batch from current inputs**

```powershell
.\生成未来三天排班.ps1 -Changes "无新增请假"
```

Expected: three `.xls` daily files, one report, one JSON summary, and no partial files.

- [ ] **Step 6: Visually compare template and all three outputs**

Open a local comparison page containing the template and three output PNGs. Confirm that only dates, personnel, and same-day notes differ, and that the table is not clipped.

- [ ] **Step 7: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/tests/test_excel_integration.py .workbuddy/skills/duty-roster-generator/scripts/export_roster_preview.ps1
git commit -m "test: verify real Excel roster copies"
```

### Task 9: Update User Workflow, Skill and Documentation

**Files:**
- Modify: `.workbuddy/skills/duty-roster-generator/SKILL.md`
- Modify: `README.md`
- Modify: `docs/使用教程.md`
- Modify: `docs/配置说明.md`
- Modify: `docs/故障排查.md`
- Modify: `安装.ps1`

- [ ] **Step 1: Document the Microsoft Excel requirement and exact-copy behavior**

State that formal daily rosters use Excel-native `.xls` copies, while the warning report continues to use the bundled spreadsheet runtime.

- [ ] **Step 2: Add installer validation**

`安装.ps1` must instantiate Excel COM, print the detected version, quit Excel, and fail with an actionable Chinese message if unavailable.

- [ ] **Step 3: Update low-token instructions**

Tell Codex to read only `运行摘要.json` after a successful run and only inspect format-difference JSON when verification fails.

- [ ] **Step 4: Validate PowerShell and skill metadata**

```powershell
$null = [scriptblock]::Create((Get-Content -Raw -Encoding UTF8 '.\安装.ps1'))
$null = [scriptblock]::Create((Get-Content -Raw -Encoding UTF8 '.\.workbuddy\skills\duty-roster-generator\scripts\write_template_roster.ps1'))
$env:PYTHONUTF8='1'
python C:\Users\26326\.codex\skills\.system\skill-creator\scripts\quick_validate.py .\.workbuddy\skills\duty-roster-generator
```

- [ ] **Step 5: Run all tests and the real one-command workflow again**

- [ ] **Step 6: Commit**

```powershell
git add .workbuddy/skills/duty-roster-generator/SKILL.md README.md docs/使用教程.md docs/配置说明.md docs/故障排查.md 安装.ps1
git commit -m "docs: document exact template roster workflow"
```

### Task 10: Release Gate and GitHub Publication

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `VERSION`

- [ ] **Step 1: Run fresh verification**

Run the full Python suite, all PowerShell syntax checks, config schema validation, skill validation, privacy scan, real three-day generation, format fingerprint comparison, and visual preview inspection.

- [ ] **Step 2: Run `/review` against the complete diff**

Fix all verified defects and rerun affected tests.

- [ ] **Step 3: Update version and changelog**

Record the template-preserving `.xls` output, transactional publication, format fingerprint verification, and Excel prerequisite.

- [ ] **Step 4: Confirm private business files remain ignored**

```powershell
git check-ignore -v -- 输入/人员/人员.xls 输入/过往三天/8.13值班表.xls 输出/*
git status --short
```

- [ ] **Step 5: Push without force**

```powershell
git push -u origin main
```

- [ ] **Step 6: Verify the remote**

Confirm the GitHub repository contains source, skill, documentation and empty input/output placeholders, but no real personnel, history or generated roster files.
