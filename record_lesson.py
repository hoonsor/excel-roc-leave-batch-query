import sys
import datetime

sys.stdout.reconfigure(encoding='utf-8')
today = datetime.datetime.now().strftime('%Y-%m-%d')
lesson = f'''
## 📌 教訓: VBA 差假查詢忽略「送審中」狀態導致漏抓

**錯誤分類**: 邏輯
**嚴重等級**: 🔴 高頻
**Token 浪費模式**: 無
**預估浪費 Tokens**: 1000

### 症狀
使用者回報「差假大批查詢工具」在比對特定教職員的公假時，未能成功在「差假資料庫」中將該筆資料標記出來，即使查詢時間區間完全重疊。

### 根因分析
經過追蹤 `modQuery.bas` 的演算法，發現在從「差假資料庫」載入資料時，狀態過濾條件僅包含 `"已簽核" Or "審核完成" Or ""`。而該筆紀錄為 **`送審中`**，導致未被載入。

### ❌ 無效嘗試（避免重蹈覆轍）
無。

### ✅ 正確解法
將 `送審中` 加入允許的狀態名單中：
```vba
If strStatus = "已簽核" Or strStatus = "審核完成" Or strStatus = "送審中" Or strStatus = "" Then
```

### 🛡️ 預防措施
處理行政機關表單與請假系統資料時，必須先遍歷一次該欄位存在的所有「狀態」枚舉值（Enumeration），再決定過濾邏輯。

### 📎 關聯
- 對話 ID: 32c05ddd-8e67-4b6f-a42c-cfad45eb41ed
- 日期: {today}
'''

with open(r'C:\Users\User\.gemini\antigravity\knowledge\hoonsor-error-learning\artifacts\error_lessons.md', 'a', encoding='utf-8') as f:
    f.write(lesson)
print('Error lesson recorded.')
