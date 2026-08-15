# -*- coding: utf-8 -*-
"""Python(win32com) 版 write_template：复制空白模板为未来每一天，
仅改写白名单单元格的值，并把模板完整格式(字体/对齐/数字格式/填充/边框/行高/列宽)逐格拷回。
比 PowerShell 版更易调试且可计时。"""
import os, sys, json, shutil, time, tempfile
import win32com.client as wc


def copy_format(src, dst):
    df = dst.Font
    sf = src.Font
    df.Name = sf.Name
    df.Size = sf.Size
    df.Bold = sf.Bold
    df.Italic = sf.Italic
    df.Underline = sf.Underline
    df.Color = sf.Color
    dst.HorizontalAlignment = src.HorizontalAlignment
    dst.VerticalAlignment = src.VerticalAlignment
    dst.NumberFormat = src.NumberFormat
    dst.Interior.ColorIndex = src.Interior.ColorIndex
    dst.Interior.Pattern = src.Interior.Pattern
    sb = src.Borders
    db = dst.Borders
    for i in (7, 8, 9, 10):  # left/top/right/bottom
        try:
            db.Item(i).LineStyle = sb.Item(i).LineStyle
            db.Item(i).Weight = sb.Item(i).Weight
            db.Item(i).ColorIndex = sb.Item(i).ColorIndex
        except Exception:
            pass


def main(payload_path):
    t0 = time.time()
    with open(payload_path, encoding="utf-8") as f:
        payload = json.load(f)
    excel = wc.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.Calculation = -4135  # xlCalculationManual
    tpl_path = os.path.abspath(payload["template_path"])
    sheet_name = payload["sheet"]
    out_dir = os.path.abspath(payload["output_dir"])
    staging = tempfile.mkdtemp(prefix="roster_staging_")
    GRID_R = 30
    GRID_C = 10
    t1 = time.time()
    for day in payload["days"]:
        target = os.path.join(staging, day["filename"])
        # 只读打开模板并另存为生成新文件，避免文件系统级复制被模板锁冲突
        tpl_book = excel.Workbooks.Open(tpl_path, ReadOnly=True)
        tpl_ws = tpl_book.Worksheets.Item(sheet_name)
        tpl_book.SaveAs(target)
        tpl_book.Close(False)
        book = excel.Workbooks.Open(target)
        ws = book.Worksheets.Item(sheet_name)
        ws.Range(day["date_cell"]).Value2 = day["date_text"]
        for c in day["cells"]:
            ws.Range(c["cell"]).Value2 = c["text"]
        ws.Range(day["note_cell"]).Value2 = day["note_text"]
        for r in range(1, GRID_R + 1):
            ws.Rows.Item(r).RowHeight = tpl_ws.Rows.Item(r).RowHeight
            for c in range(1, GRID_C + 1):
                copy_format(tpl_ws.Cells.Item(r, c), ws.Cells.Item(r, c))
        for c in range(1, GRID_C + 1):
            ws.Columns.Item(c).ColumnWidth = tpl_ws.Columns.Item(c).ColumnWidth
        book.Save()
        book.Close(False)
    t2 = time.time()
    for day in payload["days"]:
        src = os.path.join(staging, day["filename"])
        dst = os.path.join(out_dir, day["filename"])
        shutil.copy2(src, dst)
    try:
        shutil.rmtree(staging, ignore_errors=True)
    except Exception:
        pass
    t3 = time.time()
    print("INIT_MS=%.0f WRITE_MS=%.0f PUBLISH_MS=%.0f TOTAL_MS=%.0f" %
          ((t1 - t0) * 1000, (t2 - t1) * 1000, (t3 - t2) * 1000, (t3 - t0) * 1000))
    print("OK")


if __name__ == "__main__":
    main(sys.argv[1])
