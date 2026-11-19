"""
#請假查詢 工作流程腳本
用法：python scratch/leave_query_workflow.py

功能：
1. 複製「差假大批查詢工具.xlsm」並加上今日民國日期前綴，但絕對不更動、不修改、不刪除「教職員名單」工作表中的資料！
2. 使用 winsdk Windows Media OCR 讀取「002-待作業檔案」中所有 PDF（支援掃描型/圖片型 PDF）
3. 解析姓名、日期、時間並填入複製後的 xlsm 查詢介面
"""
import sys, io, os, re, shutil, subprocess, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from datetime import datetime, date

# 嘗試載入 winsdk 模組
try:
    from winsdk.windows.storage import StorageFile
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.globalization import Language
except ImportError:
    print("❌ 缺少 winsdk 模組，請先執行 pip install winsdk 安裝。")
    sys.exit(1)

import pypdfium2 as pdfium

# ── 路徑設定 ──────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    # 若為打包後的 exe 執行環境
    PROJECT_DIR = os.path.dirname(sys.executable)
else:
    # 原始開發環境
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    PROJECT_DIR = os.path.dirname(SCRIPT_DIR)

TEMPLATE     = os.path.join(PROJECT_DIR, "差假大批查詢工具.xlsm")
PDF_INBOX    = os.path.join(PROJECT_DIR, "002-待作業檔案")

ROC_OFFSET   = 1911

# ── 載入教職員名單 ─────────────────────────────────────────
def load_staff_names() -> list[str]:
    import openpyxl
    names = []
    try:
        wb = openpyxl.load_workbook(TEMPLATE, read_only=True)
        if "教職員名單" in wb.sheetnames:
            ws = wb["教職員名單"]
            for r in range(1, 1000):
                val = ws.cell(r, 1).value
                if val:
                    val_str = str(val).strip()
                    if val_str and val_str != "姓名":
                        names.append(val_str)
    except Exception as e:
        print("⚠️ 無法從範本載入教職員名單：", e)
    return names

# ── 民國日期工具 ──────────────────────────────────────────
def today_roc_str() -> str:
    today = date.today()
    roc_year = today.year - ROC_OFFSET
    return f"{roc_year:03d}{today.month:02d}{today.day:02d}"

def resolve_roc_year(school_year_str: str, month: int) -> int:
    """根據台灣學年度與月份推算民國年 (8月至翌年7月為同一學年度)"""
    try:
        school_year = int(school_year_str)
        return school_year if month >= 8 else school_year + 1
    except ValueError:
        return date.today().year - ROC_OFFSET

# ── 本地 Windows OCR 核心邏輯 ────────────────────────────────
async def ocr_page_async(png_path: str) -> str:
    abs_path = os.path.abspath(png_path)
    file = await StorageFile.get_file_from_path_async(abs_path)
    stream = await file.open_read_async()
    decoder = await BitmapDecoder.create_async(stream)
    software_bitmap = await decoder.get_software_bitmap_async()
    
    lang = Language("zh-Hant-TW")
    if OcrEngine.is_language_supported(lang):
        engine = OcrEngine.try_create_from_language(lang)
    else:
        engine = OcrEngine.try_create_from_user_profile_languages()
        
    result = await engine.recognize_async(software_bitmap)
    lines_info = []
    for line in result.lines:
        if line.words:
            y = sum(w.bounding_rect.y for w in line.words) / len(line.words)
            x = line.words[0].bounding_rect.x
            lines_info.append({'text': line.text, 'y': y, 'x': x})
            
    lines_info.sort(key=lambda item: item['y'])
    rows = []
    current_row = []
    current_y = None
    for item in lines_info:
        if current_y is None or abs(item['y'] - current_y) < 20:
            current_row.append(item)
            current_y = item['y'] if current_y is None else (current_y + item['y'])/2
        else:
            current_row.sort(key=lambda item: item['x'])
            rows.append(''.join(i['text'] for i in current_row))
            current_row = [item]
            current_y = item['y']
            
    if current_row:
        current_row.sort(key=lambda item: item['x'])
        rows.append(''.join(i['text'] for i in current_row))
        
    return '\n'.join(rows)

async def ocr_pdf_file_async(pdf_path: str) -> str:
    doc = pdfium.PdfDocument(pdf_path)
    full_text = []
    print(f"  📖 正在對 {os.path.basename(pdf_path)} 進行本機 OCR 辨識 (頁數: {len(doc)})...")
    
    for idx, page in enumerate(doc):
        bitmap = page.render(scale=3) # 解析度 3 倍提升 OCR 準確度
        pil_img = bitmap.to_pil()
        temp_png = f"temp_ocr_page_{idx}.png"
        pil_img.save(temp_png)
        
        try:
            text = await ocr_page_async(temp_png)
            full_text.append(text)
        finally:
            if os.path.exists(temp_png):
                os.remove(temp_png)
                
    return "\n".join(full_text)

# ── 空間與文字標準化 ───────────────────────────────────────
def normalize_text(text: str) -> str:
    """消除中文與數字間的多餘空格，保留分行資訊方便逐行解析"""
    lines = []
    for line in text.splitlines():
        line_clean = re.sub(r'[ \t]+', '', line)
        if line_clean:
            lines.append(line_clean)
    return "\n".join(lines)

# ── 解析提取邏輯 ─────────────────────────────────────────
def parse_extracted_data(text: str, filename: str, staff_names: list[str]) -> list[dict]:
    rows = []
    normalized = normalize_text(text)
    
    # 1. 識別學年度 (用於推算沒有寫年份的月日，如 3/12)
    school_year = "114" # 預設 fallback
    m_sy = re.search(r'(\d{3})學年度', normalized)
    if m_sy:
        school_year = m_sy.group(1)
        
    # 2. 識別姓名 (優先以檔名匹配，次以教職員名單匹配)
    name = None
    for staff in staff_names:
        if staff in filename:
            name = staff
            break
            
    # 3. 根據文件類型執行特定解析
    # 類型 A: 橫式清冊 (具備年份、月份，如 115年丙級專案... 一月份)
    m_yr = re.search(r'(\d{2,3})年', normalized)
    m_mo = re.search(r'([一二三四五六七八九十]+|\d{1,2})月份', normalized)
    if m_yr and m_mo:
        print(f"  ℹ️  使用橫式清冊(日期逗號分隔)策略...")
        yr = int(m_yr.group(1))
        mo_str = m_mo.group(1)
        chinese_num = {'一':1, '二':2, '三':3, '四':4, '五':5, '六':6, '七':7, '八':8, '九':9, '十':10, '十一':11, '十二':12}
        mo = chinese_num.get(mo_str) or int(mo_str)
        
        lines = normalized.splitlines()
        for i, line in enumerate(lines):
            matched_staff = next((s for s in staff_names if s in line), None)
            if matched_staff:
                combined_line = line
                if i + 1 < len(lines) and re.match(r'^[0-9,上下]+$', lines[i+1]):
                    combined_line += lines[i+1]
                
                if '匯款' in combined_line:
                    date_part = combined_line.split('匯款')[-1]
                else:
                    m_seq = re.search(r'([0-9,上下]+)$', combined_line)
                    date_part = m_seq.group(1) if m_seq else ''
                    
                valid_dates = re.findall(r'(?<!\d)([1-9]|[12]\d|3[01])([上下]?)(?!\d)', date_part)
                for d, part in valid_dates:
                    roc_date = f'{yr:03d}{mo:02d}{int(d):02d}'
                    start, end = '0800', '1700'
                    if part == '上': end = '1200'
                    elif part == '下': start = '1300'
                    rows.append({'name': matched_staff, 'date_roc': f"'{roc_date}", 'start': f"'{start}", 'end': f"'{end}"})
                    
        if rows:
            print(f"  📅 橫式清冊：共提取出 {len(rows)} 筆查詢時段")
            return rows

    # 類型 A2: 全國技能檢定術科工作人員清冊 (具備年份，日期為 9/22(全)、10/06(全) 格式)
    if m_yr and re.search(r'\d{1,2}/\d{1,2}\([全上下]', normalized):
        print(f"  ℹ️  使用全國檢定清冊(MM/DD(全/上/下))策略...")
        yr = int(m_yr.group(1))
        for line in normalized.splitlines():
            matched_staff = next((s for s in staff_names if s in line), None)
            if matched_staff:
                for m_d in re.finditer(r'(\d{1,2})/(\d{1,2})\((全|上|下|上午|下午)\)', line):
                    mo, d, part = int(m_d.group(1)), int(m_d.group(2)), m_d.group(3)
                    roc_date = f'{yr:03d}{mo:02d}{d:02d}'
                    start, end = '0800', '1700'
                    if part in ('上', '上午'): end = '1200'
                    elif part in ('下', '下午'): start = '1300'
                    rows.append({'name': matched_staff, 'date_roc': f"'{roc_date}", 'start': f"'{start}", 'end': f"'{end}"})
        if rows:
            print(f"  📅 全國檢定清冊：共提取出 {len(rows)} 筆查詢時段")
            return rows

    # 類型 B: 職場巡輔交通費申請表 (格式多為 3/12、4/2、5/21)
    if "職場巡輔" in normalized or "巡輔" in normalized or "交通費" in normalized:
        if not name:
            for staff in staff_names:
                if staff in normalized:
                    name = staff
                    break
        
        if not name:
            print(f"  ⚠️  無法在檔名或內文中匹配到教職員名單中的姓名，跳過此檔案")
            return []
            
        print(f"  👤 姓名：{name} (職場巡輔申請表)")
        
        seen_dates = set()
        date_patterns = [
            r'(\d{1,2})/(\d{1,2})',
            r'(\d{1,2})月(\d{1,2})日'
        ]
        
        for pat in date_patterns:
            for m in re.finditer(pat, normalized):
                try:
                    mo = int(m.group(1))
                    d = int(m.group(2))
                    if mo < 1 or mo > 12 or d < 1 or d > 31:
                        continue
                    
                    y = resolve_roc_year(school_year, mo)
                    roc_date = f"{y:03d}{mo:02d}{d:02d}"
                    
                    if roc_date not in seen_dates:
                        seen_dates.add(roc_date)
                        rows.append({"name": name, "date_roc": f"'{roc_date}", "start": "'0800", "end": "1200"})
                        rows.append({"name": name, "date_roc": f"'{roc_date}", "start": "1300", "end": "1700"})
                except Exception:
                    pass
                    
        rows = sorted(rows, key=lambda x: (x["date_roc"], x["start"]))
        print(f"  📅 職場巡輔：提取出 {len(seen_dates)} 個日期，共 {len(rows)} 筆查詢時段")
        
    # 類型 B: 印領清冊/鐘點費明細 (逐列或整頁比對日期與教師)
    else:
        print("  ℹ️  使用清冊比對策略...")
        # 3.1 找出整個 PDF 中出現過哪些教師的姓名
        teachers_in_doc = []
        for staff in staff_names:
            if staff in normalized:
                teachers_in_doc.append(staff)
        
        # 3.2 如果在整個文件中只匹配到唯一一位教師，則將文件中所有找到的日期都歸屬於他
        if len(teachers_in_doc) == 1:
            name = teachers_in_doc[0]
            m_dates = re.finditer(r'(\d{2,3})年(\d{1,2})月(\d{1,2})日', normalized)
            seen_dates = set()
            for m in m_dates:
                try:
                    y = int(m.group(1))
                    mo = int(m.group(2))
                    d = int(m.group(3))
                    roc_date = f"{y:03d}{mo:02d}{d:02d}"
                    if roc_date not in seen_dates:
                        seen_dates.add(roc_date)
                        rows.append({"name": name, "date_roc": f"'{roc_date}", "start": "'0800", "end": "'1500"})
                        print(f"  👤 匹配教師：{name} (唯一教師) | 📅 授課日期：{roc_date} (預設 0800 ~ 1500)")
                except Exception:
                    pass
        else:
            # 如果有多個教師或沒有，則使用嚴格的「同一行」匹配日期與姓名
            lines = normalized.splitlines()
            for line in lines:
                m_date = re.search(r'(\d{2,3})年(\d{1,2})月(\d{1,2})日', line)
                if m_date:
                    try:
                        y = int(m_date.group(1))
                        mo = int(m_date.group(2))
                        d = int(m_date.group(3))
                        roc_date = f"{y:03d}{mo:02d}{d:02d}"
                        
                        matched_teacher = None
                        for staff in staff_names:
                            if staff in line:
                                matched_teacher = staff
                                break
                                
                        if matched_teacher:
                            rows.append({"name": matched_teacher, "date_roc": f"'{roc_date}", "start": "'0800", "end": "'1500"})
                            print(f"  👤 匹配教師：{matched_teacher} | 📅 授課日期：{roc_date} (預設 0800 ~ 1500)")
                    except Exception:
                        pass
                        
        if not rows:
            print(f"  ⚠️  無法在清冊中匹配到任何日期與教師資料列")
            
    return rows

# ── 尋找最新差假紀錄查詢檔案 ─────────────────────────────────────
def get_latest_leave_file(project_dir: str):
    import glob
    candidates = []
    search_dirs = [project_dir, os.path.join(project_dir, "004-系統匯出之差假紀錄")]
    for sdir in search_dirs:
        if os.path.exists(sdir):
            for ext in ("*.xls", "*.xlsx", "*.xlsm"):
                for fpath in glob.glob(os.path.join(sdir, f"差假紀錄查詢_{ext}")):
                    fname = os.path.basename(fpath)
                    m = re.search(r"差假紀錄查詢_(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})", fname)
                    date_key = m.group(1) if m else ""
                    mtime = os.path.getmtime(fpath)
                    candidates.append((date_key, mtime, fpath))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]

def import_leave_records(excel, wb_dest, leave_file_path: str) -> int:
    print(f"
📥 正在自動匯入最新差假紀錄：{os.path.basename(leave_file_path)}...")
    wb_src = excel.Workbooks.Open(leave_file_path, ReadOnly=True)
    try:
        ws_src = wb_src.Sheets(1)
        last_row_src = ws_src.Cells(ws_src.Rows.Count, 1).End(-4162).Row
        if last_row_src < 2:
            print("  ⚠️ 差假紀錄檔案中無資料列")
            return 0
        raw_data = ws_src.Range(f"A2:L{last_row_src}").Value
    finally:
        wb_src.Close(False)

    valid_rows = []
    for r in raw_data:
        cell_a = str(r[0] or "").strip()
        cell_c = str(r[2] or "").strip()
        cell_e = str(r[4] or "").strip()
        if cell_a.startswith("■") or "■" in cell_a or "共計" in cell_a:
            continue
        if not cell_a and not cell_c and not cell_e:
            continue
        valid_rows.append(r)

    ws_db = wb_dest.Sheets("差假資料庫")
    ws_query = wb_dest.Sheets("查詢介面")

    if ws_db.FilterMode:
        ws_db.ShowAllData()
    if ws_db.AutoFilterMode:
        ws_db.AutoFilterMode = False

    last_row_dest = ws_db.Cells(ws_db.Rows.Count, 1).End(-4162).Row
    if last_row_dest > 1:
        ws_db.Range(f"A2:L{last_row_dest}").ClearContents()
        ws_db.Range(f"A2:L{last_row_dest}").Interior.ColorIndex = -4142
        ws_db.Range(f"A2:L{last_row_dest}").Font.ColorIndex = -4105

    if valid_rows:
        dest_range = ws_db.Range(ws_db.Cells(2, 1), ws_db.Cells(1 + len(valid_rows), 12))
        dest_range.Value = valid_rows

    fname = os.path.basename(leave_file_path)
    ws_db.Range("M1").Value = fname
    ws_query.Range("F3").Value = f"目前資料庫：{fname} ({len(valid_rows)} 筆)"
    ws_query.Range("F3").Font.Name = "微軟正黑體"
    ws_query.Range("F3").Font.Size = 10
    ws_query.Range("F3").Font.Color = 0x794E1F
    ws_query.Range("F3").Font.Bold = True

    print(f"✅ 成功匯入 {len(valid_rows)} 筆差假紀錄到「差假資料庫」！")
    return len(valid_rows)

# ── Excel 寫入 ────────────────────────────────────────────
def write_to_excel(xlsm_path: str, rows: list[dict], leave_file_path: str = None) -> int:
    import win32com.client
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    imported_count = 0
    
    try:
        wb = excel.Workbooks.Open(xlsm_path)
        
        # 1. 若有指定最新差假紀錄檔案，先行自動匯入差假資料庫
        if leave_file_path and os.path.exists(leave_file_path):
            imported_count = import_leave_records(excel, wb, leave_file_path)
        
        print(f"
✍️ 正在將 PDF 查詢時段寫入 Excel 查詢介面...")
        ws_query = wb.Sheets("查詢介面")
    import win32com.client
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        print(f"\n✍️ 正在將資料寫入 Excel...")
        wb = excel.Workbooks.Open(xlsm_path)
        ws_query = wb.Sheets("查詢介面")
        
        # 1. 清除「查詢介面」舊有查詢資料 (B5 以後)，但絕對不碰其他工作表！
        last_used = ws_query.Cells(ws_query.Rows.Count, 2).End(-4162).Row # xlUp = -4162
        if last_used >= 5:
            ws_query.Range(f"A5:G{last_used}").ClearContents()
            ws_query.Range(f"A5:G{last_used}").Interior.ColorIndex = -4142 # xlNone
            ws_query.Range(f"A5:G{last_used}").Font.ColorIndex = -4105      # xlAutomatic
            
        # 2. 寫入新提取的查詢資料
        for i, row in enumerate(rows):
            r = 5 + i
            ws_query.Cells(r, 1).Value = i + 1              # 序號
            ws_query.Cells(r, 2).Value = row["name"]         # 姓名
            ws_query.Cells(r, 3).NumberFormat = "@"
            ws_query.Cells(r, 3).Value = row["date_roc"]     # 日期 (字串)
            ws_query.Cells(r, 4).NumberFormat = "@"
            ws_query.Cells(r, 4).Value = row["start"]        # 開始時間
            ws_query.Cells(r, 5).NumberFormat = "@"
            ws_query.Cells(r, 5).Value = row["end"]          # 結束時間
            
        wb.Save()
        print(f"🎉 成功寫入 {len(rows)} 筆資料到：{os.path.basename(xlsm_path)}")
        
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        # 強制關閉 Excel 背景進程釋放檔案鎖
        subprocess.run(
            ["powershell", "-Command",
             "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
            capture_output=True
        )

# ── 主流程 ────────────────────────────────────────────────
async def main_async():
    print("=" * 60)
    print("  #請假查詢 工作流程啟動 (winsdk OCR 離線版)")
    print("=" * 60)
    
    # 0. 載入教職員名單
    staff_names = load_staff_names()
    print(f"👥 已載入 {len(staff_names)} 位教職員名單作比對基礎")
    if not staff_names:
        print("❌ 教職員名單為空，無法進行比對識別，請先在「差假大批查詢工具.xlsm」的教職員名單分頁中建置名單。")
        return
        
    # 1. 取得今日民國日期
    roc_today = today_roc_str()
    print(f"📅 今日民國日期：{roc_today}")
    
    # 2. 複製活頁簿 (原地保留所有工作表與名單)
    if not os.path.exists(TEMPLATE):
        print(f"❌ 找不到範本檔案：{TEMPLATE}")
        return
        
    dest_name = f"{roc_today}-差假大批查詢工具.xlsm"
    dest_path = os.path.join(PROJECT_DIR, dest_name)
    
    # 如果已存在，直接複製覆蓋，這會帶有原範本所有的工作表格式
    shutil.copy2(TEMPLATE, dest_path)
    print(f"✅ 成功建立今日查詢工作簿：{dest_name}")
    
    # 3. 掃描待作業 PDF
    if not os.path.exists(PDF_INBOX):
        print(f"❌ 找不到待作業資料夾：{PDF_INBOX}")
        return
        
    pdf_files = [
        os.path.join(PDF_INBOX, f)
        for f in sorted(os.listdir(PDF_INBOX))
        if f.lower().endswith(".pdf")
    ]
    
    if not pdf_files:
        print(f"\n📂 「002-待作業檔案」資料夾目前沒有 PDF 檔案。")
        print("   請放檔案後再執行 #請假查詢。")
        return
        
    print(f"\n📂 找到 {len(pdf_files)} 個待處理 PDF 檔案：")
    for p in pdf_files:
        print(f"  - {os.path.basename(p)}")
        
    # 4. 對每個 PDF 進行 OCR 辨識與解析
    all_rows = []
    for pdf_path in pdf_files:
        try:
            text = await ocr_pdf_file_async(pdf_path)
            rows = parse_extracted_data(text, os.path.basename(pdf_path), staff_names)
            all_rows.extend(rows)
        except Exception as e:
            print(f"  ❌ {os.path.basename(pdf_path)} 辨識或解析失敗：{e}")
            import traceback
            traceback.print_exc()
            
    if not all_rows:
        print("\n⚠️ 檔案解析後無有效時段資料，請確認 PDF 內容。")
        return
        
    print(f"\n📊 統計：共從 {len(pdf_files)} 個 PDF 中提取出 {len(all_rows)} 筆查詢時段")
    
    # 5. 寫入複製的 Excel 檔案中
    # 5. 自動尋找目錄下最新之差假紀錄檔
    latest_leave_file = get_latest_leave_file(PROJECT_DIR)
    if latest_leave_file:
        print(f"\n📂 找到最新差假紀錄檔：{os.path.basename(latest_leave_file)}")
    else:
        print("\n⚠️ 未在目錄下找到差假紀錄查詢檔案，將保留範本既有之資料庫。")
        
    # 6. 寫入複製的 Excel 檔案中 (同時自動匯入差假紀錄與寫入查詢資料)
    imported_count = write_to_excel(dest_path, all_rows, latest_leave_file)
    
    print("\n" + "=" * 60)
    print("  🎉 #請假查詢 任務已完成！")
    print("=" * 60)
    print(f"\n  📁 今日工作檔案：{dest_name}")
    if latest_leave_file:
        print(f"  📥 已自動匯入差假紀錄：{os.path.basename(latest_leave_file)} (共 {imported_count} 筆)")
    print(f"  📄 已辨識 PDF 數：{len(pdf_files)} 個")
    print(f"  📝 填入查詢筆數：{len(all_rows)} 筆")
    print("\n  💡 教職員名單保護：")
    print("     您的「教職員名單」工作表完全完好如初，絕無任何刪除或修改。")
    print("\n  ➡️  後續操作：")
    print("     1. 請開啟檔案")
    if latest_leave_file:
        print("     2. 系統已自動幫您匯入最新差假紀錄，您只需直接點選「2. 執行大批查詢」即可開始比對！")
    else:
        print("     2. 點選「1. 匯入請假紀錄檔」匯入差勤 XLS 檔")
        print("     3. 點選「2. 執行大批查詢」進行比對與著色")

if __name__ == "__main__":
    try:
        asyncio.run(main_async())
    except Exception as e:
        print(f"\n❌ 發生未預期錯誤：{e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n✅ 任務結束")
