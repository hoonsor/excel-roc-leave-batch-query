import os
import re
from datetime import datetime
import subprocess

def update_preferences():
    config_dir = r"C:\Users\User\.gemini\config"
    pref_path = os.path.join(config_dir, "skills", "hoonsor-preferences", "PREFERENCES.md")
    status_path = os.path.join(config_dir, "PROJECT_STATUS.md")
    
    today = datetime.now().strftime("%Y-%m-%d")
    
    # -------------------------------------------------------------
    # 1. 讀取並更新 PREFERENCES.md
    # -------------------------------------------------------------
    with open(pref_path, 'r', encoding='utf-8') as f:
        pref_text = f.read()
        
    excel_section = """## 📊 Excel VBA 與巨集工具開發偏好 [新增於 2026-07-05]
- **一鍵式整合工作流**：在主畫面（如「查詢介面」）提供直覺的按鈕控制流程（如：1. 匯入、2. 執行比對、3. 清除結果），減少跨工作表的來回切換與操作複雜度。
- **高效記憶體運作**：大數據量比對時，必須一次性載入記憶體陣列或使用 Dictionary 雜湊字典進行高速索引，嚴禁巢狀迴圈讀寫儲存格，比對速度應在 1 秒內完成。
- **雙向視覺同步與標記**：
  - 對於查詢匹配到的請假資料，除了在查詢介面標註（淡紅色背景與紅色字體），應同步將原始「差假資料庫」工作表中對應的原始資料列 (row) 標記上相同顏色。
  - 比對完成後，自動啟用背景色自動篩選 (AutoFilter by Cell Color)，方便使用者在資料庫中一眼確認是哪幾筆紀錄被命中。
- **無損復原與預防干擾**：
  - **文字防呆**：日期或時間欄位（如民國日期 YYYMMDD、起迄時間 HHMM）必須設定或格式化為文字格式（`@`），防止 Excel 自動吞掉 `0800` 等前導字元 `0`。
  - **清除機制**：清除查詢時同步清空資料庫的著色、取消篩選，且必須**保留資料庫內容，僅還原視圖與著色**。
  - **新查詢預清空**：每次新查詢前，自動預先清空前一次資料庫的篩選與著色，防止舊顏色標記干擾新比對。
"""
    
    if "Excel VBA 與巨集工具開發偏好" not in pref_text:
        # 在 "---"（倒數第二個）或末尾前追加
        # 我們直接在 "---" 前方插入
        parts = pref_text.split("---")
        if len(parts) >= 2:
            parts[-2] = parts[-2].strip() + "\n\n" + excel_section + "\n"
            pref_text = "---".join(parts)
        else:
            pref_text = pref_text.strip() + "\n\n" + excel_section
            
        with open(pref_path, 'w', encoding='utf-8') as f:
            f.write(pref_text)
        print("Successfully updated PREFERENCES.md")
    else:
        print("PREFERENCES.md already contains Excel section, skipping.")
        
    # -------------------------------------------------------------
    # 2. 讀取並更新 PROJECT_STATUS.md (全域)
    # -------------------------------------------------------------
    with open(status_path, 'r', encoding='utf-8') as f:
        status_text = f.read()
        
    if "v1.6.5" not in status_text:
        # 1) 更新版本號：> **版本號：** `v1.6.4` -> `v1.6.5`
        status_text = re.sub(r"> \*\*版本號：\*\* `v1.6.4`", "> **版本號：** `v1.6.5`", status_text)
        # 2) 更新最後更新：> **最後更新：** 2026-06-18 -> 2026-07-05
        status_text = re.sub(r"> \*\*最後更新：\*\* 2026-06-18", f"> **最後更新：** {today}", status_text)
        
        # 3) 插入版本歷程表格行
        new_row = f"| v1.6.5 | {today} | feat(pref) | 透過 #喜好 指令自動將「Excel 直覺化與人性化設計規範」錄入 PREFERENCES.md 偏好設定庫 |"
        target_row_marker = "| v1.6.4 | 2026-06-18 | fix | 重構同步腳本、技能引導與還原指南中的硬編碼路徑，改用動態環境變數以支援異機還原無縫運作 |"
        if target_row_marker in status_text:
            status_text = status_text.replace(target_row_marker, f"{new_row}\n{target_row_marker}")
            
        # 4) 在主要任務追加 Checked Checkbox
        new_task = f"- [x] 透過 #喜好 自動將「Excel 直覺化與人性化設計規範」錄入 PREFERENCES.md 偏好庫 (v1.6.5)"
        task_marker = "- [x] 重構同步腳本、技能引導與還原指南中的硬編碼路徑，改用動態環境變數以支援異機還原無縫運作 (v1.6.4)"
        if task_marker in status_text:
            status_text = status_text.replace(task_marker, f"{new_task}\n{task_marker}")
            
        with open(status_path, 'w', encoding='utf-8') as f:
            f.write(status_text)
        print("Successfully updated global PROJECT_STATUS.md")
    else:
        print("global PROJECT_STATUS.md already contains v1.6.5, skipping.")
        
    # -------------------------------------------------------------
    # 3. 執行 Git 的 add, commit (不在此處強推送，若有權限問題則在 cli 反饋)
    # -------------------------------------------------------------
    try:
        # 在 C:\Users\User\.gemini\config 目錄下執行
        subprocess.run(["git", "add", "-A"], cwd=config_dir, check=True)
        subprocess.run(["git", "commit", "-m", "feat(pref): add Excel usability guidelines to preferences library"], cwd=config_dir, check=True)
        print("Successfully committed changes to global Git repo.")
        
        # 嘗試推送
        res = subprocess.run(["git", "push"], cwd=config_dir, capture_output=True, text=True)
        if res.returncode == 0:
            print("Successfully pushed changes to remote GitHub repo.")
        else:
            print(f"Git push failed (possibly due to authentication/network issues): {res.stderr.strip()}")
    except Exception as e:
        print(f"Git operation error: {e}")

if __name__ == "__main__":
    update_preferences()
