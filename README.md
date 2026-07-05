# Excel 巨集與民國曆日期計算工具箱 (Excel Macro & ROC Date Toolbox)

本專案是一個專門為台灣行政、財務及公務體系設計的 Excel VBA 巨集工具箱，主要解決 Excel 原生不支援中華民國曆（民國年）日期計算與轉換的問題。

---

## 🌟 核心特色
*   **專業角色**：本專案配置有專屬的 `AGENTS.md` 規則，使 AI 助理自動扮演 Excel 巨集大師及台灣民國曆日期計算專家。
*   **安全轉換**：严格遵循民國年與西元年轉換邏輯，杜絕魔術數字，使用 `ROC_YEAR_OFFSET = 1911` 常數進行運算。
*   **閏年相容**：完美判斷民國曆的閏年（與西元閏年規則連動），精準計算 2 月天數。
*   **格式健全**：解析各種常見格式如 `YYYMMDD` (7 位數)、`YYMMDD` (6 位數，如 990520) 以及 `YYY/MM/DD`。

---

## 📁 目錄結構說明
*   `src/`：VBA 原始碼模組（`.bas`、`.cls`、`.frm`）。
*   `tests/`：單元測試與驗證模組。
*   `.agents/`：專案 AI 協同規則檔。
*   `ANTIGRAVITY.md`：專案最高指導原則。
*   `PROJECT_STATUS.md`：專案進度與版本歷程。

---

## 🚀 快速開始
1.  閱讀 [ANTIGRAVITY.md](file:///d:/%E5%8C%97%E7%A7%91%E9%99%84%E5%B7%A5/015-Antigravity%E5%B7%A5%E4%BD%9C%E8%B3%87%E6%96%99%E5%A4%BE/03-Antigravity%E5%8D%94%E5%8A%A9%E6%92%B0%E5%AF%AB%E5%B7%A8%E9%9B%86/ANTIGRAVITY.md) 瞭解技術規範。
2.  檢視 [.agents/AGENTS.md](file:///d:/%E5%8C%97%E7%A7%91%E9%99%84%E5%B7%A5/015-Antigravity%E5%B7%A5%E4%BD%9C%E8%B3%87%E6%96%99%E5%A4%BE/03-Antigravity%E5%8D%94%E5%8A%A9%E6%92%B0%E5%AF%AB%E5%B7%A8%E9%9B%86/.agents/AGENTS.md) 瞭解中華民國曆計算標準。
