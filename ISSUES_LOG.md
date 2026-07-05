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

---

## 🛑 [Issue #002] VBA 編譯錯誤：只有定義於公用物件模組中的使用者自訂型態，才可以和 variant 型態互轉

### 1. 問題描述
解決編碼問題後，執行查詢時在 `colLeaves.Add rec` 行發生編譯錯誤。提示 `只有定義於公用物件模組中的使用者自訂型態，才可以和 variant 型態互轉，也才能傳遞給延遲連結(late-bound)的函數。`

### 2. 根本原因 (Root Cause)
*   在 `modQuery.bas` 中，`colLeaves` 是 VBA 的 `Collection`。
*   VBA 的 `Collection.Add` 方法接受的參數型態為 `Variant`。
*   VBA 限制：在標準模組中以 `Private Type` 定義的**自訂結構 (User-Defined Type, UDT)** 不能自動轉換成 `Variant`，因此無法直接加入 `Collection` 或 `Dictionary` 中儲存。

### 3. 解決方案 (Solution)
*   **改用類別模組**：建立一個新的 VBA 類別模組 [clsLeaveRecord.cls](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/src/clsLeaveRecord.cls)，將原先 `LeaveRecord` UDT 宣告的欄位轉換為此類別的 `Public` 屬性。
*   **程式重構**：重構 [modQuery.bas](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/src/modQuery.bas)，將 `LeaveRecord` 宣告替換為 `clsLeaveRecord`，並在存入 Collection 前使用 `Set rec = New clsLeaveRecord` 進行實例化。
*   **更新建置管線**：在 [build_and_test.py](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/scratch/build_and_test.py) 與 [fix_encoding.py](file:///d:/北科附工/015-Antigravity工作資料夾/03-Antigravity協助撰寫巨集/scratch/fix_encoding.py) 中，將 `clsLeaveRecord.cls` 加入轉碼與匯入清單。

### 4. 預防計畫 (Prevention Plan)
*   在 VBA 開發中，若需要將結構化資料存入 `Collection` 或 `Dictionary`，應**優先設計類別模組 (Class Module) 作為資料載體**，避免使用 `Type` 導致與 `Variant` 不相容的限制。
