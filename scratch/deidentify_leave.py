"""
去識別化腳本：將差假紀錄查詢 XLS 中的「姓名」欄位替換為 mock 資料。
- 此檔案無身份證字號欄位，只有姓名 (Col 2)
- 同一個真實姓名 → 對應到同一個 mock 姓名（一致性保持）
- 輸出新檔案，原始檔案不被覆蓋
"""
import sys, io, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import xlrd
from xlutils.copy import copy as xl_copy

SRC_PATH = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集\差假紀錄查詢_2026-07-05-09-50-38.xls"
DST_PATH = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集\差假紀錄查詢_去識別化.xls"

COL_NAME = 2  # 姓名欄位索引

def gen_mock_name(index: int) -> str:
    """產生無意義的 mock 姓名（格式：員工 + 3位數編號）"""
    return f"員工{index + 1:03d}"

def main():
    print(f"Reading source: {SRC_PATH}")
    rb = xlrd.open_workbook(SRC_PATH, formatting_info=True)
    ws_r = rb.sheet_by_index(0)

    print(f"Total rows: {ws_r.nrows}, Cols: {ws_r.ncols}")

    # 第一遍：蒐集所有唯一姓名，按首次出現順序建立映射
    real_to_mock = {}   # real_name -> mock_name
    mock_counter = 0

    for r in range(1, ws_r.nrows):
        real_name = str(ws_r.cell_value(r, COL_NAME)).strip()
        if real_name and real_name not in real_to_mock:
            mock_name = gen_mock_name(mock_counter)
            real_to_mock[real_name] = mock_name
            mock_counter += 1

    print(f"Unique names found: {mock_counter}")
    print("Mapping sample (first 10):")
    for i, (real, mock) in enumerate(list(real_to_mock.items())[:10]):
        print(f"  [{real}] → [{mock}]")
    if mock_counter > 10:
        print(f"  ... and {mock_counter - 10} more")

    # 第二遍：複製並替換
    wb_w = xl_copy(rb)
    ws_w = wb_w.get_sheet(0)

    replaced_rows = 0
    for r in range(1, ws_r.nrows):
        real_name = str(ws_r.cell_value(r, COL_NAME)).strip()
        if real_name in real_to_mock:
            ws_w.write(r, COL_NAME, real_to_mock[real_name])
            replaced_rows += 1

    wb_w.save(DST_PATH)
    print(f"\n✅ De-identified file saved: {DST_PATH}")
    print(f"   Rows replaced: {replaced_rows}")

    # 驗證：確認原始姓名完全不在輸出中
    print("\nRunning verification...")
    rb2 = xlrd.open_workbook(DST_PATH)
    ws2 = rb2.sheet_by_index(0)
    original_names = set(real_to_mock.keys())
    found_originals = []
    for r in range(1, ws2.nrows):
        val = str(ws2.cell_value(r, COL_NAME)).strip()
        if val in original_names:
            found_originals.append((r, val))

    if found_originals:
        print(f"⚠️  WARNING: {len(found_originals)} original names still found in output!")
        for row, val in found_originals[:5]:
            print(f"   Row {row}: {val}")
    else:
        print(f"🔒 Verification PASSED: No original names found in output ({ws2.nrows} rows checked).")

    # 顯示輸出前幾行供確認
    print("\n=== Output sample (rows 1-4) ===")
    for r in range(1, min(5, ws2.nrows)):
        name_v = ws2.cell_value(r, COL_NAME)
        unit_v = ws2.cell_value(r, 0)
        start_v = ws2.cell_value(r, 4)
        print(f"  Row {r}: Name=[{name_v}] Unit=[{unit_v}] Start=[{start_v}]")

if __name__ == "__main__":
    main()
