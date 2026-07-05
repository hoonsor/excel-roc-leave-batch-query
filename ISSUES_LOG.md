# 專案問題紀錄簿 (Issues Log)

本文件記錄了本專案中遇到並解決的技術問題、根本原因、解決方案及預防措施。

---

## 🛑 [Issue #001] VBA 程式碼導入後發生「編譯錯誤：語法錯誤」

### 1. 問題描述
使用者點擊「匯入請假明細」按鈕時，Excel 彈出 `編譯錯誤：語法錯誤`。開啟 VBE 視窗發現所有中文註解、工作表名稱與字串提示皆顯示為亂碼，且部分 VBA 語法關鍵字（如 `Public Sub`、`Dim`）字元遭到截斷或損毀。

### 2. 根本原因 (Root Cause)
*   **編碼衝突**：VBA 編輯器 (VBE) 在執行 `.VBProject.VBComponents.Import` 匯入 `.bas` 模組檔案時，預設會使用作業系統的 ANSI 字元集（在台灣繁體中文 Windows 環境下為 **CP950 / Big5**）進行檔案解碼。
*   **來源格式不符**：原本由程式開發產出的 `.bas` 原始碼檔案是採用 **UTF-8** 編碼儲存。當 VBE 使用 Big5 強行解碼 UTF-8 的中文字元時，由於位元組解析錯誤，甚至將某些 UTF-8 位元組判定為換行、退格或控制字元，進而損毀了前後的 VBA 語法結構，導致無法編譯。

### 3. 解決方案 (Solution)
*   **轉碼處理**：撰寫轉碼腳本 [fix_encoding.py](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/scratch/fix_encoding.py)，將 `src/` 底下的三個主要 VBA 檔案（`modUtility.bas`、`modImport.bas`、`modQuery.bas`）從 UTF-8 重新編碼為 **CP950 (Big5)** 覆蓋儲存。
*   **防止鎖定重構**：更新 [build_and_test.py](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/scratch/build_and_test.py)，加入檔案佔用偵測與強殺背景殘留 Excel 進程的功能。
*   **重新構建**：重置編碼後，重新以 CP950 原始碼檔重新生成並包裝 `差假大批查詢工具.xlsm`。

### 4. 預防計畫 (Prevention Plan)
*   **VBA 檔案強制 ANSI 儲存**：未來所有提供給 Excel 匯入使用的原始碼檔案（如 `.bas`、`.cls` 等），在建立與編輯時皆必須強制以 CP950 / ANSI 編碼進行儲存。
*   **防呆檢驗**：在建置指令中保留自動化測試功能，確保每一次代碼異動都會在背景試跑 Excel 測試以檢驗是否發生 compile error。
