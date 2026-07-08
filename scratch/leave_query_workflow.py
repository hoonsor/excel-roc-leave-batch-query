"""
#請假查詢 工作流程腳本
用法：python scratch/leave_query_workflow.py

功能：
1. 複製「差假大批查詢工具.xlsm」並加上今日民國日期前綴
2. 讀取「002-待作業檔案」中所有 PDF
3. 解析姓名、日期、時間並填入複製後的 xlsm 查詢介面
"""
import sys, io, os, re, shutil, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from datetime import datetime, date

# ── 路徑設定 ──────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR  = os.path.dirname(SCRIPT_DIR)
TEMPLATE     = os.path.join(PROJECT_DIR, "差假大批查詢工具.xlsm")
PDF_INBOX    = os.path.join(PROJECT_DIR, "002-待作業檔案")

ROC_OFFSET   = 1911  # 民國紀年偏移

# ── 民國日期工具 ──────────────────────────────────────────
def today_roc_str() -> str:
    """回傳今日民國日期字串，格式 YYYMMDD（7碼）。"""
    today = date.today()
    roc_year = today.year - ROC_OFFSET
    return f"{roc_year:03d}{today.month:02d}{today.day:02d}"

def parse_roc_date(text: str) -> str | None:
    """
    從文字中嘗試解析民國日期，傳回 YYYMMDD 格式字串。
    支援格式：115/07/08、115-07-08、115年7月8日、1150708
    """
    text = text.strip()
    # 7碼純數字
    m = re.fullmatch(r'(\d{3})(\d{2})(\d{2})', text)
    if m:
        return text
    # YYY/MM/DD 或 YYY-MM-DD
    m = re.search(r'(\d{3})[/-](\d{1,2})[/-](\d{1,2})', text)
    if m:
        y, mo, d = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
        return f"{y}{mo}{d}"
    # YYY年M月D日
    m = re.search(r'(\d{2,3})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
    if m:
        y  = m.group(1).zfill(3)
        mo = m.group(2).zfill(2)
        d  = m.group(3).zfill(2)
        return f"{y}{mo}{d}"
    return None

# ── PDF 讀取 ──────────────────────────────────────────────
def read_pdf_text(pdf_path: str) -> str:
    """讀取 PDF 全文。優先使用 pdfplumber，失敗則嘗試 PyMuPDF。"""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except ImportError:
        pass
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        return "\n".join(page.get_text() for page in doc)
    except ImportError:
        raise RuntimeError("請安裝 pdfplumber 或 PyMuPDF：pip install pdfplumber 或 pip install pymupdf")

# ── PDF 解析策略 ──────────────────────────────────────────
def parse_patrol_pdf(text: str, filename: str) -> list[dict]:
    """
    解析「職場巡輔交通費申請表」PDF。
    回傳 list of dict: [{name, date_roc, start, end}, ...]
    """
    rows = []
    lines = text.splitlines()

    # 從檔名取得姓名（格式：...-申請表-姓名.pdf）
    name_from_filename = None
    m = re.search(r'申請表-(.+?)\.pdf', filename, re.IGNORECASE)
    if m:
        name_from_filename = m.group(1).strip()

    # 從內文找「申請人」或「姓名」
    name = name_from_filename
    for line in lines:
        m = re.search(r'申請人[：:]\s*([^\s]+)', line)
        if m:
            name = m.group(1).strip()
            break
        m = re.search(r'姓\s*名[：:]\s*([^\s]+)', line)
        if m:
            name = m.group(1).strip()
            break

    if not name:
        print(f"  ⚠️  無法從 {filename} 辨識姓名，跳過")
        return []

    print(f"  👤 姓名：{name}")

    # 找所有日期（過濾重複）
    found_dates = []
    seen = set()
    for line in lines:
        # 尋找民國日期格式
        for m in re.finditer(r'\d{2,3}[年/\-]\s*\d{1,2}[月/\-]\s*\d{1,2}[日]?', line):
            roc = parse_roc_date(m.group())
            if roc and roc not in seen and len(roc) == 7:
                seen.add(roc)
                found_dates.append(roc)

    print(f"  📅 找到 {len(found_dates)} 個巡輔日期")

    # 每個日期預設拆為上下午兩筆
    for d in found_dates:
        rows.append({"name": name, "date_roc": d, "start": "0800", "end": "1200"})
        rows.append({"name": name, "date_roc": d, "start": "1300", "end": "1700"})

    return rows

def parse_generic_pdf(text: str, filename: str) -> list[dict]:
    """通用解析：嘗試從 PDF 文字中提取姓名與日期。"""
    print(f"  ℹ️  使用通用解析策略處理：{filename}")
    rows = []
    lines = text.splitlines()

    name = None
    for line in lines:
        m = re.search(r'姓名[：:]\s*([^\s,，]+)', line)
        if m:
            name = m.group(1).strip()
            break

    if not name:
        # 嘗試從檔名取得
        m = re.search(r'[\-_]([^\-_\.]+)\.pdf$', filename)
        if m:
            name = m.group(1)

    if not name:
        print(f"  ⚠️  無法辨識姓名，跳過：{filename}")
        return []

    seen = set()
    for line in lines:
        for m in re.finditer(r'\d{2,3}[年/\-]\s*\d{1,2}[月/\-]\s*\d{1,2}[日]?', line):
            roc = parse_roc_date(m.group())
            if roc and roc not in seen:
                seen.add(roc)
                rows.append({"name": name, "date_roc": roc, "start": "0800", "end": "1700"})

    return rows

def parse_pdf(pdf_path: str) -> list[dict]:
    """根據 PDF 檔名或內容選擇對應的解析策略。"""
    filename = os.path.basename(pdf_path)
    print(f"\n📄 處理：{filename}")
    text = read_pdf_text(pdf_path)

    if "巡輔" in filename or "巡輔" in text[:500]:
        return parse_patrol_pdf(text, filename)
    else:
        return parse_generic_pdf(text, filename)

# ── Excel 填寫 ────────────────────────────────────────────
def write_to_excel(xlsm_path: str, rows: list[dict]):
    """將解析結果填入 xlsm 查詢介面（從第 5 列開始）。"""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(xlsm_path)
        ws = wb.Sheets("查詢介面")

        # 清除舊資料（B5 以後）
        last_used = ws.Cells(ws.Rows.Count, 2).End(-4162).Row  # xlUp = -4162
        if last_used >= 5:
            ws.Range(f"A5:G{last_used}").ClearContents()
            ws.Range(f"A5:G{last_used}").Interior.ColorIndex = -4142
            ws.Range(f"A5:G{last_used}").Font.ColorIndex = -4105

        # 填入新資料
        for i, row in enumerate(rows):
            r = 5 + i
            ws.Cells(r, 1).Value = i + 1            # 序號
            ws.Cells(r, 2).Value = row["name"]       # 姓名
            # 日期欄設為文字格式，防止前導零丟失
            ws.Cells(r, 3).NumberFormat = "@"
            ws.Cells(r, 3).Value = row["date_roc"]   # 日期 YYYMMDD
            ws.Cells(r, 4).NumberFormat = "@"
            ws.Cells(r, 4).Value = row["start"]      # 開始時間
            ws.Cells(r, 5).NumberFormat = "@"
            ws.Cells(r, 5).Value = row["end"]        # 結束時間

        wb.Save()
        print(f"\n✅ 已填入 {len(rows)} 筆查詢資料至：{xlsm_path}")

    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        # 強制釋放檔案鎖定
        subprocess.run(
            ["powershell", "-Command",
             "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
            capture_output=True
        )

# ── 主流程 ────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  #請假查詢 工作流程啟動")
    print("=" * 60)

    # Step 1: 今日民國日期
    roc_today = today_roc_str()
    print(f"\n📅 今日民國日期：{roc_today}")

    # Step 2: 複製工具檔案
    if not os.path.exists(TEMPLATE):
        print(f"❌ 找不到模板檔案：{TEMPLATE}")
        sys.exit(1)

    dest_name = f"{roc_today}-差假大批查詢工具.xlsm"
    dest_path = os.path.join(PROJECT_DIR, dest_name)

    if os.path.exists(dest_path):
        print(f"⚠️  目的地已存在：{dest_name}，將直接覆蓋。")

    shutil.copy2(TEMPLATE, dest_path)
    print(f"✅ 已複製：{dest_name}")

    # Step 3: 掃描 PDF
    if not os.path.exists(PDF_INBOX):
        print(f"❌ 找不到待作業資料夾：{PDF_INBOX}")
        sys.exit(1)

    pdf_files = [
        os.path.join(PDF_INBOX, f)
        for f in sorted(os.listdir(PDF_INBOX))
        if f.lower().endswith(".pdf")
    ]

    if not pdf_files:
        print(f"\n📂 「002-待作業檔案」資料夾目前沒有 PDF 檔案。")
        print("   請將需要查詢的 PDF 放入後再執行 #請假查詢。")
        sys.exit(0)

    print(f"\n📂 找到 {len(pdf_files)} 個 PDF：")
    for p in pdf_files:
        print(f"  - {os.path.basename(p)}")

    # Step 4: 解析所有 PDF
    all_rows = []
    for pdf_path in pdf_files:
        try:
            rows = parse_pdf(pdf_path)
            all_rows.extend(rows)
            print(f"  → 產生 {len(rows)} 筆查詢資料")
        except Exception as e:
            print(f"  ❌ 解析失敗：{e}")

    if not all_rows:
        print("\n⚠️  所有 PDF 解析後無有效資料，請確認 PDF 格式。")
        sys.exit(1)

    print(f"\n📊 合計 {len(all_rows)} 筆查詢資料（來自 {len(pdf_files)} 個 PDF）")

    # Step 5: 填入 Excel
    write_to_excel(dest_path, all_rows)

    # Step 6: 完成通知
    print("\n" + "=" * 60)
    print("  🎉 #請假查詢 工作流程完成！")
    print("=" * 60)
    print(f"\n  📁 輸出檔案：{dest_name}")
    print(f"  📄 已處理 PDF：{len(pdf_files)} 個")
    print(f"  📝 填入查詢筆數：{len(all_rows)} 筆")
    print("\n  ➡️  接下來請：")
    print("  1. 開啟輸出的 xlsm 檔案")
    print("  2. 點選「1. 匯入請假明細」匯入差假資料庫")
    print("  3. 點選「2. 執行大批查詢」開始比對")

if __name__ == "__main__":
    main()
