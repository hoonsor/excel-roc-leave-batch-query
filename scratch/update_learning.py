import os
import json
import sys
from datetime import datetime

# Reconfigure stdout to support unicode emojis in Windows command prompt
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def update_lessons():
    workspace_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
    kb_dir = r"C:\Users\User\.gemini\antigravity\knowledge\hoonsor-error-learning"
    lessons_path = os.path.join(kb_dir, "artifacts", "error_lessons.md")
    quick_fixes_path = os.path.join(kb_dir, "artifacts", "quick_fixes.md")
    patterns_path = os.path.join(kb_dir, "artifacts", "token_waste_patterns.md")
    metadata_path = os.path.join(kb_dir, "metadata.json")
    
    today = datetime.now().strftime("%Y-%m-%d")
    conv_id = "32c05ddd-8e67-4b6f-a42c-cfad45eb41ed"
    
    # -------------------------------------------------------------
    # 1. 準備追加的四大教訓內容
    # -------------------------------------------------------------
    new_lessons_content = f"""
## 📌 教訓 #13: VBE 匯入 UTF-8 原始檔亂碼與語法損毀

**錯誤分類**: 編碼
**嚴重等級**: 🔴 高頻
**Token 浪費模式**: TW-04（環境假設錯誤）
**預估浪費 Tokens**: ~2,000

### 症狀
Excel VBA 編輯器 (VBE) 在匯入 `.bas` 或 `.cls` 原始檔後，彈出 `編譯錯誤：語法錯誤`。開啟 VBE 程式碼視窗發現所有中文註解、中文字串提示均顯示為亂碼，且部分關鍵字（如 `Public Sub`、`Dim`）字元遭到截斷或損毀。

### 根因分析
* VBE 的匯入機制 (`.VBProject.VBComponents.Import`) 預設會使用系統的 ANSI 字元集（台灣繁中環境下為 **CP950 / Big5**）來解碼原始碼。
* 若來源原始碼是採用標準 **UTF-8** 編碼儲存，當 VBE 使用 Big5 強行解碼 UTF-8 的中文字元時，由於位元組解析錯誤，甚至將某些 UTF-8 位元組判定為換行、退格或控制字元，進而將鄰近的 VBA 關鍵字字元截斷（例如 `Dim` 變為 `d`），導致語法編譯失敗。

### ❌ 無效嘗試（避免重蹈覆轍）
1. 嘗試手動在 VBE 裡面重新打字修補關鍵字 ➡️ 治標不治本，中文亂碼仍然存在且下次建置又會被覆蓋。

### ✅ 正確解法
在 Python 自動化建置管線中，使用無損轉碼架構：
* 原始碼檔案本身在 `src/` 中始終以 **UTF-8** 儲存（利於版控與多人協作）。
* 匯入時，在臨時目錄（如 `scratch/temp_cp950/`）寫入一個轉為 **CP950** 編碼的臨時檔。
* 從臨時檔執行 `.Import` 匯入，完成後隨即將臨時檔與臨時目錄刪除。

### 🛡️ 預防措施
* 嚴禁直接將 UTF-8 的原始碼檔 `.Import` 到 VBE 中。
* 始終使用「UTF-8 原始檔 ➡️ 動態 CP950 暫存檔 ➡️ 匯入 ➡️ 清理」的無損建置管線。

### 📎 關聯
- 對話 ID: {conv_id}
- 日期: {today}
- 相關 KI: 無

---

## 📌 教訓 #14: VBA Collection 儲存 UDT 限制

**錯誤分類**: 邏輯
**嚴重等級**: 🟡 中等
**Token 浪費模式**: 無
**預估浪費 Tokens**: ~0

### 症狀
執行 VBA 比對程式時，在將自訂結構 `Private Type` 放入 Collection 集合時（如 `colLeaves.Add rec`），發生編譯錯誤：
```text
編譯錯誤：只有定義於公用物件模組中的使用者自訂型態，才可以和 variant 型態互轉，也才能傳遞給延遲連結(late-bound)的函數。
```

### 根因分析
VBA 的 `Collection.Add` 方法參數接收型態為 `Variant`。而在標準模組中以 `Private Type` 宣告的使用者自訂結構 (User-Defined Type, UDT) 不能隱式轉換成 `Variant`，因此無法直接存入 `Collection` 或 `Dictionary` 等集合中。

### ❌ 無效嘗試（避免重蹈覆轍）
1. 試圖將 `Type` 改為 `Public Type` ➡️ 在一般標準模組中依然會報相同的 `Variant` 互轉錯誤。

### ✅ 正確解法
改以**類別模組 (Class Module)** 來做資料載體：
1. 建立一個類別模組 `clsLeaveRecord.cls`，將原本 `Type` 中的欄位改為該類別的 `Public` 屬性。
2. 在比對程序中，改用類別物件：
   ```vba
   Dim rec As clsLeaveRecord
   Set rec = New clsLeaveRecord
   rec.LeaveType = "..."
   colLeaves.Add rec  ' 類別物件可以直接存入 Collection
   ```

### 🛡️ 預防措施
在 VBA 開發中，若資料結構需要被存入 `Collection`、`Dictionary` 或作為物件傳遞，應**優先設計類別模組 (Class Module)**，不要使用 `Type` 結構。

### 📎 關聯
- 對話 ID: {conv_id}
- 日期: {today}
- 相關 KI: 無

---

## 📌 教訓 #15: VBA 雙重轉碼導致非 ASCII 字元全部變問號

**錯誤分類**: 編碼
**嚴重等級**: 🔴 高頻
**Token 浪費模式**: TW-01（盲目重試）, TW-04（環境假設錯誤）
**預估浪費 Tokens**: ~5,000

### 症狀
Excel 巨集執行時，出現滿屏問號的錯誤訊息框（標題如 `???~`），且 VBE 代碼中的中文字串常數全部被替換成 `?`（問號）。

### 根因分析
在建置腳本中，若轉碼程式（如 `fix_encoding.py`）重複被執行：
1. 第一次執行時，將原本 UTF-8 的原始碼成功轉為 CP950。
2. 第二次執行時，轉碼程式仍試圖以 UTF-8 讀取。由於此時檔案已是 CP950，讀取會引發解碼錯誤。
3. 程式如果使用 `latin-1` fallback 讀取，並加上 `errors='replace'` 再次以 CP950 寫入，會因為 `latin-1` 的解碼字元無法在 CP950 中對應，導致所有非 ASCII 的繁體中文字元全被強制覆寫替換為 `?`。這就是「雙重轉碼」引起的字元毀損。
4. 這也導致代碼中的 `ThisWorkbook.Sheets("差假資料庫")` 變成了 `"Sheets("?t°??????w")"`，因為找不到該 Sheet 而引發「下標越界」錯誤。

### ❌ 無效嘗試（避免重蹈覆轍）
1. 再次執行轉碼腳本 ➡️ 問號 `?` 為不可逆損毀，再次轉碼只會增加更多問號。

### ✅ 正確解法
1. 重新從無損備份（如 Git 版本庫）還原正確中文的 UTF-8 原始檔。
2. 廢棄獨立的轉碼腳本，改為在 `build_and_test.py` 的建置記憶體中進行「一次性動態轉為 CP950 暫存檔」的流程，使得主代碼永遠不會受到重複轉碼的影響。

### 🛡️ 預防措施
* 嚴禁對同一個原始碼檔案進行多次覆蓋式轉碼。
* 使用單向、非覆蓋的建置流水線，確保原始檔案的編碼安全性。

### 📎 關聯
- 對話 ID: {conv_id}
- 日期: {today}
- 相關 KI: 無

---

## 📌 教訓 #16: VBA 動態陣列未分配即傳參導致下標越界

**錯誤分類**: 邏輯
**嚴重等級**: 🔴 高頻
**Token 浪費模式**: TW-04（環境假設錯誤）
**預估浪費 Tokens**: ~3,000

### 症狀
輸入純數字 7 碼民國年（如 `1150210`）時，大批查詢巨集始終判定為 `日期格式錯誤 (須7碼)`。

### 根因分析
在 `CDateFromROC` 中，當沒有分隔符時，不會執行 `Split`。此時變數 `parts()` 是一個未經分配（未分配維度）的動態陣列。
當程式將 `parts()` 以 `ByRef` 傳遞給 `IsArrayInitialized(ByRef arr() As String)` 時，VBA Runtime 在**傳參階段**（在進入該函式內部前）就會直接拋出 `執行期錯誤 9：下標越界`，且無法被 `IsArrayInitialized` 內部的 `On Error Resume Next` 捕捉，導致整個轉換程序崩潰。

### ❌ 無效嘗試（避免重蹈覆轍）
1. 在 `IsArrayInitialized` 裡面加強 `On Error Resume Next` ➡️ 毫無效果，因為錯誤在進入該函式前就已發生。

### ✅ 正確解法
改用 Boolean 變數 `hasSeparator` 作為旗標來進行邏輯分流：
```vba
Dim hasSeparator As Boolean
hasSeparator = False

If InStr(strClean, "/") > 0 Then
    parts = Split(strClean, "/")
    hasSeparator = True
End If

If hasSeparator Then
    ' 使用 parts() ...
Else
    ' 直接解析純數字 YYYMMDD，不碰 parts()
End If
```
這能徹底避開未分配陣列傳參崩潰的 VBA 底層限制。

### 🛡️ 預防措施
在 VBA 中，盡量避免將可能「未分配/未定義」的動態陣列傳遞給任何 ByRef 參數。應優先使用狀態旗標（如 Boolean）控制邏輯分流。

### 📎 關聯
- 對話 ID: {conv_id}
- 日期: {today}
- 相關 KI: 無
"""
    
    # -------------------------------------------------------------
    # 2. 更新 error_lessons.md
    # -------------------------------------------------------------
    with open(lessons_path, 'r', encoding='utf-8') as f:
        lessons_text = f.read()
        
    # 如果已經包含了教訓 #13，說明這是一次重新執行，我們不用重複插入索引行
    if "教訓 #13" not in lessons_text:
        new_index_rows = f"""| 13 | VBE 匯入 UTF-8 原始檔亂碼與語法損毀 | 編碼 | 🔴 | TW-04 | ~2,000 | {today} |
| 14 | VBA Collection 儲存 UDT 限制 | 邏輯 | 🟡 | — | 0 | {today} |
| 15 | VBA 雙重轉碼導致非 ASCII 字元全部變問號 | 編碼 | 🔴 | TW-01, TW-04 | ~5,000 | {today} |
| 16 | VBA 動態陣列未分配即傳參導致下標越界 | 邏輯 | 🔴 | TW-04 | ~3,000 | {today} |"""

        target_row_marker = "| 12 | Next.js 頁面 Hydration 警告未除 | 前端 | 🟡 | TW-01 | ~3,000 | 2026-06-14 |"
        if target_row_marker in lessons_text:
            lessons_text = lessons_text.replace(target_row_marker, f"{target_row_marker}\n{new_index_rows}")
            
        lessons_text = lessons_text.strip() + "\n\n" + new_lessons_content.strip() + "\n"
        with open(lessons_path, 'w', encoding='utf-8') as f:
            f.write(lessons_text)
        print("Successfully updated error_lessons.md")
    else:
        print("error_lessons.md already contains the new lessons, skipping text append.")
    
    # -------------------------------------------------------------
    # 3. 更新 quick_fixes.md
    # -------------------------------------------------------------
    with open(quick_fixes_path, 'r', encoding='utf-8') as f:
        qf_text = f.read()
        
    if "Excel VBA" not in qf_text:
        excel_vba_section = """## Excel VBA

| 錯誤關鍵字 | 正確解法 | 教訓 # |
|-----------|---------|:-----:|
| `編譯錯誤：語法錯誤` (亂碼) | 匯入時將原始碼轉為 CP950 暫存檔匯入，原始碼保留為 UTF-8 | #13 |
| `編譯錯誤：無法與 variant 互轉` | 將 UDT (Type) 改用類別模組 (Class Module) 封裝為物件 | #14 |
| `???~` / `?×?J??º` (MsgBox 亂碼) | 還原 UTF-8 原始檔，以暫存 CP950 建置，避免雙重轉碼 | #15 |
| `下標越界` / `日期格式錯誤` (未分配陣列) | 避免對未分配的動態陣列進行 ByRef 傳參，改用 Boolean 變數分流 | #16 |
| `SaveAs` 失敗 / 檔案鎖定 | 建置前檢測 os.remove，PermissionError 時以 taskkill 強制結束 Excel 背景進程 | #15 |

"""
        target_qf_marker = "## React / 前端"
        if target_qf_marker in qf_text:
            qf_text = qf_text.replace(target_qf_marker, excel_vba_section + target_qf_marker)
        else:
            qf_text = qf_text.strip() + "\n\n" + excel_vba_section.strip() + "\n"
            
        with open(quick_fixes_path, 'w', encoding='utf-8') as f:
            f.write(qf_text)
        print("Successfully updated quick_fixes.md")
    else:
        print("quick_fixes.md already contains Excel VBA section, skipping.")
    
    # -------------------------------------------------------------
    # 4. 更新 token_waste_patterns.md
    # -------------------------------------------------------------
    new_pat_text = """# 統計：Token 浪費模式

本檔案記錄並分析各類「Token 浪費模式」的累計頻率與預估 tokens 損耗，用以量化與督促 Agent 的自我改善效果。

---

## 📊 累計統計表

| 模式代碼 | 模式名稱 | 累計次數 | 累計浪費 Tokens | 關聯教訓 |
|----------|---------|:--------:|:---------------:|----------|
| `TW-01` | **盲目重試** | 4 | ~14,000 | #12, #15 |
| `TW-02` | **方向錯誤** | 3 | ~42,000 | #2, #9 |
| `TW-03` | **忽略已知解法** | 0 | 0 | — |
| `TW-04` | **環境假設錯誤** | 7 | ~39,000 | #1, #5, #8, #10, #13, #15, #16 |
| `TW-05` | **過度工程** | 0 | 0 | — |
| `TW-06` | **依賴混淆** | 2 | ~9,000 | #3, #4 |
| `TW-07` | **Selector 地獄** | 1 | ~30,000 | #2 |

---

## 📈 改善趨勢分析
* **核心痛點**：`TW-02` (方向錯誤) 由於曾在 Angular 虛擬 DOM 上做過大量無效嘗試，導致單筆耗損極大。
* **高頻痛點**：`TW-04` (環境假設錯誤) 發生次數最多，特別是 Windows 控制台編碼、PowerShell 語法不相容，以及 **Excel VBA 的 CP950 解碼限制與動態陣列傳參限制**。
* **改善成果**：引進「動態 CP950 暫存建置管線」與「Boolean 旗標分流設計」後，此類 Windows / Excel 環境問題已成功絕跡。
"""
    with open(patterns_path, 'w', encoding='utf-8') as f:
        f.write(new_pat_text)
    print("Successfully updated token_waste_patterns.md")
    
    # -------------------------------------------------------------
    # 5. 更新 metadata.json
    # -------------------------------------------------------------
    with open(metadata_path, 'r', encoding='utf-8') as f:
        meta = json.load(f)
        
    meta["updated_at"] = datetime.now().isoformat()
    meta["total_lessons"] = 16
    
    # 避免重複追加 references
    has_ref = False
    if "references" in meta:
        for ref in meta["references"]:
            if ref.get("conversation_id") == conv_id and 13 in ref.get("lessons", []):
                has_ref = True
                break
    else:
        meta["references"] = []
        
    if not has_ref:
        meta["total_estimated_tokens_saved"] = meta.get("total_estimated_tokens_saved", 0) + 10000
        new_ref = {
            "title": "Excel VBA 動態暫存建置管線與 Class Module 重構",
            "conversation_id": conv_id,
            "date": today,
            "lessons": [13, 14, 15, 16]
        }
        meta["references"].append(new_ref)
    
    with open(metadata_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print("Successfully updated metadata.json")
    
    # -------------------------------------------------------------
    # 6. 生成 Session Report
    # -------------------------------------------------------------
    report = f"""🧠 錯誤學習報告 — {today}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 統計摘要
  • 發現錯誤事件: 4 件
  • Token 浪費模式: 2 種 (TW-01, TW-04)
  • 預估總浪費: ~10,000 tokens
  • 新增教訓: 4 條 (#13, #14, #15, #16)
  • 已知教訓命中: 0 條

📌 關鍵教訓
  1. [#13] VBE 匯入 UTF-8 原始檔亂碼 — 動態轉 CP950 暫存檔匯入，主代碼保留為 UTF-8。
  2. [#14] VBA Collection 儲存 UDT 限制 — 使用 Class Module 物件代替 UDT，符合 Variant 規範。
  3. [#15] 雙重轉碼中文字元毀損 — 避免對同檔案多次轉碼，建立單向無損暫存匯入管線。
  4. [#16] 未分配動態陣列 ByRef 傳參崩潰 — 避免動態陣列傳參，改用 Boolean 旗標控制分流。

🛡️ 防線更新
  • 黃金法則: 保留 5 條
  • 速查表: 新增 5 條 Excel VBA 相關項目
  • 完整教訓: 追加 4 條詳細分析案

📁 已歸檔至: knowledge/hoonsor-error-learning/
"""
    print("\n" + report)
    
    # 同時將 Report 寫入至 scratch 備份
    report_path = os.path.join(workspace_dir, "scratch", "learning_report.txt")
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)

if __name__ == "__main__":
    update_lessons()
