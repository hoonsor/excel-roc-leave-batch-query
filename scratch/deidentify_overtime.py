"""
去識別化腳本：將加班紀錄查詢 XLS 中的「身份證字號」與「姓名」欄位
替換為 mock 資料，其他欄位維持原樣。
- 同一個真實身份證字號 → 對應到同一組 mock 資料（一致性保持）
- 輸出新檔案，原始檔案不被覆蓋
"""
import sys, io, os, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import xlrd
from xlutils.copy import copy as xl_copy

FOLDER = os.path.dirname(os.path.abspath(__file__))
# 往上一層到專案資料夾
PROJECT_DIR = os.path.dirname(FOLDER)

SRC_FILENAME = '加班紀錄查詢_2026-07-08-11-48-37.xls'
DST_FILENAME = '加班紀錄查詢_2026-07-08-11-48-37_去識別化.xls'

SRC_PATH = os.path.join(PROJECT_DIR, SRC_FILENAME)
DST_PATH = os.path.join(PROJECT_DIR, DST_FILENAME)

COL_ID   = 0  # 身份證字號
COL_NAME = 3  # 姓名

def gen_mock_id(index: int) -> str:
    """產生格式合法但純 mock 的身份證字號（英文字母 + 9 位數字）"""
    letters = "ABCDEFGHJKLMNPQRSTUVXYZ"
    prefix = letters[index % len(letters)]
    random.seed(index * 9973 + 1234)
    digits = "".join(str(random.randint(0, 9)) for _ in range(9))
    return f"{prefix}{digits}"

def gen_mock_name(index: int) -> str:
    return f"員工{index + 1:03d}"

def main():
    print(f"Source: {SRC_PATH}")
    print(f"Output: {DST_PATH}")
    print(f"Source exists: {os.path.exists(SRC_PATH)}")

    rb = xlrd.open_workbook(SRC_PATH, formatting_info=True)
    ws_r = rb.sheet_by_index(0)
    print(f"Sheet: {rb.sheet_names()[0]}, Rows: {ws_r.nrows}, Cols: {ws_r.ncols}")
    print()

    # 建立真實 ID → mock 資料的映射（一致性保持）
    real_id_to_mock = {}   # real_id -> (mock_id, mock_name)
    mock_counter = 0

    for r in range(1, ws_r.nrows):
        real_id   = str(ws_r.cell_value(r, COL_ID)).strip()
        real_name = str(ws_r.cell_value(r, COL_NAME)).strip()
        if real_id and real_id not in real_id_to_mock:
            mock_id   = gen_mock_id(mock_counter)
            mock_name = gen_mock_name(mock_counter)
            real_id_to_mock[real_id] = (mock_id, mock_name)
            print(f"  [{real_id}] [{real_name}] → [{mock_id}] [{mock_name}]")
            mock_counter += 1

    print(f"\nTotal unique persons: {mock_counter}")

    # 複製活頁簿並替換
    wb_w = xl_copy(rb)
    ws_w = wb_w.get_sheet(0)

    replaced_rows = 0
    for r in range(1, ws_r.nrows):
        real_id = str(ws_r.cell_value(r, COL_ID)).strip()
        if real_id and real_id in real_id_to_mock:
            mock_id, mock_name = real_id_to_mock[real_id]
            ws_w.write(r, COL_ID,   mock_id)
            ws_w.write(r, COL_NAME, mock_name)
            replaced_rows += 1

    wb_w.save(DST_PATH)
    print(f"\n✅ Saved: {DST_PATH}")
    print(f"   Rows replaced: {replaced_rows}")

    # 驗證
    rb2 = xlrd.open_workbook(DST_PATH)
    ws2 = rb2.sheet_by_index(0)
    original_ids = set(real_id_to_mock.keys())
    found = []
    for r in range(1, ws2.nrows):
        val = str(ws2.cell_value(r, COL_ID)).strip()
        if val in original_ids:
            found.append((r, val))

    if found:
        print(f"\n⚠️  WARNING: {len(found)} original IDs still found in output!")
    else:
        print(f"\n🔒 Verification PASSED: No original IDs remain ({ws2.nrows} rows checked).")

    # 輸出樣本
    print("\n=== Output sample (rows 1-5) ===")
    for r in range(1, min(6, ws2.nrows)):
        id_v   = ws2.cell_value(r, COL_ID)
        name_v = ws2.cell_value(r, COL_NAME)
        unit_v = ws2.cell_value(r, 1)
        type_v = ws2.cell_value(r, 4)
        if id_v or name_v:
            print(f"  Row {r}: ID=[{id_v}] Name=[{name_v}] Unit=[{unit_v}] Type=[{type_v}]")

if __name__ == "__main__":
    main()
