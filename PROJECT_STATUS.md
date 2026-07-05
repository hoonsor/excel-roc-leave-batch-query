# 專案狀態 (Project Status)

本文件用以追蹤「Excel 巨集與民國曆日期計算工具箱」專案的最新開發進度、功能清單與版本歷程。

---

## 📊 專案基本資訊
*   **專案名稱**：Excel 巨集與民國曆日期計算工具箱
*   **當前版本**：`v1.0.0`
*   **開發狀態**：已完成全部開發與自動化整合測試 (Completed & Verified)
*   **負責人**：Antigravity (Excel 巨集大師 & 民國曆日期計算專家)

---

## 🗺️ 開發地圖與功能清單 (Roadmap & Feature List)

### 1. 基礎架構與規則設定
*   [x] 建立工作區專案規則檔案 `AGENTS.md` (定義民國曆計算與 VBA 風格規範)
*   [x] 建立引導文件 `ANTIGRAVITY.md` (定義專案技術棧與目錄結構)
*   [x] 建立專案狀態文件 `PROJECT_STATUS.md`

### 2. 民國曆核心自訂函數 (ROC Calendar Core UDFs)
*   [x] 實作 `CDateFromROC`：解析 YYYMMDD 或 YYY/MM/DD 民國字串為 VBA Date
*   [x] 實作 `FormatROC`：將 VBA Date 轉換為指定格式之民國日期字串
*   [x] 實作 `IsRocLeapYear`：判斷民國年份是否為閏年
*   [x] 實作單元測試巨集並在 Excel 內執行驗證

### 3. Excel 巨集實用工具箱 (Excel Macro Utilities)
*   [x] 實作「匯入請假明細」功能：支援選擇外部請假檔匯入，排除彙總統計行
*   [x] 實作「大批請假查詢」功能：時間重疊比對（字典與記憶體陣列優化）
*   [x] 實作「清除查詢內容」功能：重設查詢介面並清除著色格式
*   [x] 完成 Python 自動化建置與測試腳本，實現一鍵整合驗證

---

## 📈 版本歷程 (Version History)

| 版本 | 日期 | 異動內容 | 負責人 |
| :--- | :--- | :--- | :--- |
| `v1.0.0` | 2026-07-05 | 完成 VBA 原始碼開發、巨集活頁簿美化與整合測試，發行正式版本 `v1.0.0`。 | Antigravity |
| `v0.1.0` | 2026-07-05 | 初始化專案，建立 `AGENTS.md`、`ANTIGRAVITY.md` 與 `PROJECT_STATUS.md`。 | Antigravity |
