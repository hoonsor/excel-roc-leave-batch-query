# 專案狀態 (Project Status)

本文件用以追蹤「Excel 巨集與民國曆日期計算工具箱」專案的最新開發進度、功能清單與版本歷程。

---

## 📊 專案基本資訊
*   **專案名稱**：Excel 巨集與民國曆日期計算工具箱
*   **當前版本**：`v0.1.0`
*   **開發狀態**：已完成專案初始化與規則設定 (Planning/Initialization)
*   **負責人**：Antigravity (Excel 巨集大師 & 民國曆日期計算專家)

---

## 🗺️ 開發地圖與功能清單 (Roadmap & Feature List)

### 1. 基礎架構與規則設定
*   [x] 建立工作區專案規則檔案 `AGENTS.md` (定義民國曆計算與 VBA 風格規範)
*   [x] 建立引導文件 `ANTIGRAVITY.md` (定義專案技術棧與目錄結構)
*   [x] 建立專案狀態文件 `PROJECT_STATUS.md`

### 2. 民國曆核心自訂函數 (ROC Calendar Core UDFs)
*   [ ] 實作 `CDateFromROC`：解析 YYYMMDD 或 YYY/MM/DD 民國字串為 VBA Date
*   [ ] 實作 `FormatROC`：將 VBA Date 轉換為指定格式之民國日期字串
*   [ ] 實作 `IsRocLeapYear`：判斷民國年份是否為閏年
*   [ ] 實作單元測試巨集並在 Excel 內執行驗證

### 3. Excel 巨集實用工具箱 (Excel Macro Utilities)
*   [ ] 實作「批次轉換儲存格日期」巨集（西元 <-> 民國）
*   [ ] 實作自動格式化民國日期儲存格工具
*   [ ] 提供 VBA 匯出與匯入公用程式 (用於原始碼管理)

---

## 📈 版本歷程 (Version History)

| 版本 | 日期 | 異動內容 | 負責人 |
| :--- | :--- | :--- | :--- |
| `v0.1.0` | 2026-07-05 | 初始化專案，建立 `AGENTS.md`、`ANTIGRAVITY.md` 與 `PROJECT_STATUS.md`。 | Antigravity |
