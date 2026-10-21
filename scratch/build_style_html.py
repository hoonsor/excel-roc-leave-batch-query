import os

base_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"

styles = [
    {
        "id": "樣式 01",
        "title": "職場巡輔交通費申請表",
        "orient": "直式表格",
        "img": "docs/style_previews/style_01_patrol.png",
        "date_fmt": "表格內含 3/12、4/2 或 X月X日（依標題學年度自動推算跨年之民國年）",
        "time_rule": "每一巡輔日期自動拆分為「上午 0800~1200」與「下午 1300~1700」兩筆查詢紀錄",
        "files": "1150708-職場巡輔交通費申請表（共 7 份歷史申請表）",
        "badge": "自動拆分上下午"
    },
    {
        "id": "樣式 02",
        "title": "丙級專案技能檢定印領清冊",
        "orient": "橫式清冊",
        "img": "docs/style_previews/style_02_level_c.png",
        "date_fmt": "右上角標註「X月份」，日期欄以逗號分隔（如 2,3,4上,19下），姓名欄自動去空白",
        "time_rule": "無標註 = 0800~1700（整天）｜標註「上」= 0800~1200｜標註「下」= 1300~1700",
        "files": "1150918-115年丙級專案技能檢定清冊（1月、3月、5月）",
        "badge": "逗號分隔 + 上/下"
    },
    {
        "id": "樣式 03",
        "title": "身心障礙學習扶助印領清冊鐘點費",
        "orient": "直式清冊",
        "img": "docs/style_previews/style_03_special_ed.png",
        "date_fmt": "手寫或打字之授課月日與節次（如 4/1、4/8、4/27）",
        "time_rule": "依清冊上標註之實際授課起迄時段提取（如 1610~1710、1610~1810、0900~1500）",
        "files": "1150708-身心障礙學習扶助印領清冊鐘點費-1 ~ -6（共 6 份清冊）",
        "badge": "手寫/特定節次時段"
    },
    {
        "id": "樣式 04",
        "title": "產業特殊需求類科教學鐘點費",
        "orient": "直式表格",
        "img": "docs/style_previews/style_04_industry.png",
        "date_fmt": "多位教師並列之表格，含具體授課年月日",
        "time_rule": "依表格各列標註之具體授課起迄時間（如 0810~1200、1610~1800、0800~1700）",
        "files": "1150708-產業特殊需求類科教學鐘點費.pdf（共 21 位教師、59 筆時段）",
        "badge": "多教師並列明細"
    },
    {
        "id": "樣式 05",
        "title": "英文科鐘點費印領清冊",
        "orient": "橫式表格",
        "img": "docs/style_previews/style_05_english.png",
        "date_fmt": "橫式印領清冊，含教師姓名與授課日期（如 6/30、7/10）",
        "time_rule": "依清冊標註之實際授課時間（如 1000~1300、0800~1200、1300~1700）",
        "files": "1150715-英文科鐘點費-1.pdf（共 3 筆時段）",
        "badge": "科別專案鐘點"
    },
    {
        "id": "樣式 06",
        "title": "聯課活動鐘點費印領清冊",
        "orient": "橫式清冊",
        "img": "docs/style_previews/style_06_club.png",
        "date_fmt": "全體教師統一活動日期（如 115年8月19日）+ 多頁教師名冊",
        "time_rule": "固定社團／聯課活動時段 1300~1500",
        "files": "1150909-115-1-8月聯課活動鐘點費印領清冊.pdf（共 69 位教師）",
        "badge": "批次名單統一時段"
    },
    {
        "id": "樣式 07",
        "title": "代理導師職務印領名冊",
        "orient": "橫式名冊",
        "img": "docs/style_previews/style_07_homeroom.png",
        "date_fmt": "包含代理起迄日期區間（如 2026/8/17 ~ 2026/8/19 或單日）",
        "time_rule": "跨日區間自動展開為「逐日單筆」紀錄，預設整天 0800~1700",
        "files": "1150911-115年8月代理導師職務印領名冊(0817~0831).pdf（共展開 44 筆）",
        "badge": "日期區間自動展開"
    },
    {
        "id": "樣式 08",
        "title": "全國檢定乙級術科工作人員酬勞費印領清冊",
        "orient": "直式表格",
        "img": "docs/style_previews/style_08_national_test.png",
        "date_fmt": "標題含「115年度」，日期欄為頓號分隔之 MM/DD(全) 格式（如 9/22(全)、10/06(全)）",
        "time_rule": "(全) = 0800~1700（整天）｜(上) = 0800~1200｜(下) = 1300~1700",
        "files": "1151019-115年度第2梯次全國檢定建築物室內設計乙級術科技術士檢定術科人員酬勞費印領清冊.pdf（共 96 筆）",
        "badge": "2026-10-21 最新驗證通過"
    }
]

cards_html = []
for s in styles:
    img_src = s["img"]
    cards_html.append(f"""
    <div class="card">
      <div class="card-header">
        <span class="style-id">{s["id"]}</span>
        <span class="orient-badge">{s["orient"]}</span>
        <span class="feature-badge">{s["badge"]}</span>
      </div>
      <h3 class="card-title">{s["title"]}</h3>
      <div class="img-wrapper" onclick="openLightbox('{img_src}', '{s["id"]} - {s["title"]}')">
        <img src="{img_src}" alt="{s["title"]}" loading="lazy" />
        <div class="zoom-hint">🔍 點擊放大檢視原圖（已去識別化）</div>
      </div>
      <div class="card-body">
        <div class="rule-row">
          <span class="rule-label">📅 日期特徵：</span>
          <span class="rule-text">{s["date_fmt"]}</span>
        </div>
        <div class="rule-row">
          <span class="rule-label">⏰ 時間轉換：</span>
          <span class="rule-text highlight">{s["time_rule"]}</span>
        </div>
        <div class="rule-row">
          <span class="rule-label">📂 歷史案例：</span>
          <span class="rule-text muted">{s["files"]}</span>
        </div>
        <div class="rule-row">
          <span class="rule-label">🖼️ 圖片路徑：</span>
          <span class="rule-text muted"><code>{img_src}</code></span>
        </div>
      </div>
    </div>
    """)

html_content = f"""<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>北科附工 - 差假大批查詢工具｜已支援 PDF 清冊樣式圖鑑</title>
  <style>
    :root {{
      --primary: #4F3314;
      --primary-light: #794E1F;
      --accent: #D97706;
      --bg: #F8FAFC;
      --card-bg: #FFFFFF;
      --text: #1E293B;
      --muted: #64748B;
      --border: #E2E8F0;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
      padding-bottom: 60px;
    }}
    header {{
      background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
      color: #fff;
      padding: 36px 24px;
      text-align: center;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
    }}
    header h1 {{ font-size: 28px; margin-bottom: 8px; letter-spacing: 1px; }}
    header p {{ font-size: 15px; opacity: 0.9; max-width: 850px; margin: 0 auto; }}
    .container {{ max-width: 1360px; margin: 28px auto; padding: 0 20px; }}
    
    .workflow-box {{
      background: #FFFBEB;
      border: 1px solid #FDE68A;
      border-left: 6px solid var(--accent);
      border-radius: 12px;
      padding: 20px 24px;
      margin-bottom: 32px;
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
    }}
    .wf-step h4 {{ color: #92400E; font-size: 16px; margin-bottom: 6px; }}
    .wf-step p {{ font-size: 14px; color: #78350F; }}
    .cmd-tag {{
      display: inline-block;
      background: #4F3314;
      color: #fff;
      padding: 2px 10px;
      border-radius: 6px;
      font-weight: bold;
      font-family: Consolas, monospace;
      margin: 0 2px;
    }}

    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(400px, 1fr));
      gap: 24px;
    }}
    .card {{
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 14px;
      overflow: hidden;
      box-shadow: 0 4px 12px rgba(0,0,0,0.05);
      transition: transform 0.2s, box-shadow 0.2s;
      display: flex;
      flex-direction: column;
    }}
    .card:hover {{
      transform: translateY(-4px);
      box-shadow: 0 12px 24px rgba(0,0,0,0.1);
    }}
    .card-header {{
      padding: 14px 18px 8px;
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .style-id {{
      background: var(--primary);
      color: #fff;
      font-weight: bold;
      font-size: 13px;
      padding: 3px 10px;
      border-radius: 99px;
    }}
    .orient-badge {{
      background: #E0F2FE;
      color: #0369A1;
      font-size: 12px;
      font-weight: bold;
      padding: 3px 10px;
      border-radius: 99px;
    }}
    .feature-badge {{
      background: #DCFCE7;
      color: #15803D;
      font-size: 12px;
      font-weight: bold;
      padding: 3px 10px;
      border-radius: 99px;
      margin-left: auto;
    }}
    .card-title {{
      padding: 0 18px 12px;
      font-size: 19px;
      color: var(--primary);
      border-bottom: 1px solid var(--border);
    }}
    .img-wrapper {{
      position: relative;
      height: 310px;
      background: #F1F5F9;
      overflow: hidden;
      cursor: zoom-in;
      display: flex;
      align-items: center;
      justify-content: center;
      border-bottom: 1px solid var(--border);
    }}
    .img-wrapper img {{
      max-width: 96%;
      max-height: 96%;
      object-fit: contain;
      box-shadow: 0 2px 8px rgba(0,0,0,0.12);
      transition: transform 0.25s;
    }}
    .img-wrapper:hover img {{ transform: scale(1.03); }}
    .zoom-hint {{
      position: absolute;
      bottom: 10px;
      right: 12px;
      background: rgba(15, 23, 42, 0.78);
      color: #fff;
      font-size: 12px;
      padding: 4px 10px;
      border-radius: 6px;
      pointer-events: none;
    }}
    .card-body {{
      padding: 16px 18px;
      display: flex;
      flex-direction: column;
      gap: 10px;
      flex: 1;
    }}
    .rule-row {{ font-size: 14px; display: flex; flex-direction: column; }}
    .rule-label {{ font-weight: bold; color: #334155; margin-bottom: 2px; }}
    .rule-text {{ color: #1E293B; }}
    .rule-text.highlight {{
      background: #FEF3C7;
      color: #92400E;
      padding: 6px 10px;
      border-radius: 6px;
      font-weight: bold;
    }}
    .rule-text.muted {{ color: var(--muted); font-size: 13px; }}
    code {{ background: #F1F5F9; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace; }}

    /* Lightbox */
    #lightbox {{
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(15, 23, 42, 0.88);
      z-index: 9999;
      align-items: center;
      justify-content: center;
      flex-direction: column;
      padding: 20px;
      cursor: zoom-out;
    }}
    #lightbox.active {{ display: flex; }}
    #lightbox img {{
      max-width: 95vw;
      max-height: 86vh;
      object-fit: contain;
      border-radius: 8px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.5);
      background: #fff;
    }}
    #lightbox-caption {{
      color: #fff;
      margin-top: 14px;
      font-size: 18px;
      font-weight: bold;
    }}
    #lightbox-close {{
      position: absolute;
      top: 20px;
      right: 28px;
      color: #fff;
      font-size: 32px;
      cursor: pointer;
      font-weight: bold;
    }}
  </style>
</head>
<body>
  <header>
    <h1>🏫 北科附工 - 差假大批查詢工具｜已支援 PDF 清冊樣式圖鑑</h1>
    <p>收錄目前已學習並驗證通過之 8 大類印領清冊與申請表樣式（所有範例圖片中之教職員人名與簽章皆已使用白色色塊隱匿去識別化，並完整保留日期與時段格式）。</p>
  </header>

  <div class="container">
    <div class="workflow-box">
      <div class="wf-step">
        <h4>🚀 步驟 1：正式請假查詢（已知樣式）</h4>
        <p>將 PDF 放入 <b>002-待作業檔案</b> 資料夾，對 AI 輸入 <span class="cmd-tag">#請假查詢</span>（建議加註「使用AI判斷模式」），系統即會依下方樣式規則自動填入今日 Excel。</p>
      </div>
      <div class="wf-step">
        <h4>🧪 步驟 2：測試新清冊樣式（未知樣式）</h4>
        <p>將新格式 PDF 放入 <b>003-測試資料</b> 資料夾，對 AI 輸入 <span class="cmd-tag">#測試學習</span> 或 <span class="cmd-tag">#測試</span>，確認無誤後即可新增至本圖鑑中。</p>
      </div>
      <div class="wf-step">
        <h4>📊 步驟 3：Excel 一鍵匯入與比對</h4>
        <p>打開產生的 <b>.xlsm</b> 檔案，點擊「<b>1. 匯入請假紀錄檔</b>」載入差勤 XLS，再點擊「<b>2. 執行大批查詢</b>」完成標色與篩選。</p>
      </div>
    </div>

    <div class="grid">
      {"".join(cards_html)}
    </div>
  </div>

  <div id="lightbox" onclick="closeLightbox()">
    <span id="lightbox-close">&times;</span>
    <img id="lightbox-img" src="" alt="Preview" />
    <div id="lightbox-caption"></div>
  </div>

  <script>
    function openLightbox(src, title) {{
      document.getElementById('lightbox-img').src = src;
      document.getElementById('lightbox-caption').textContent = title;
      document.getElementById('lightbox').classList.add('active');
    }}
    function closeLightbox() {{
      document.getElementById('lightbox').classList.remove('active');
    }}
    document.addEventListener('keydown', function(e) {{
      if (e.key === 'Escape') closeLightbox();
    }});
  </script>
</body>
</html>
"""

out_html = os.path.join(base_dir, "已處理樣式圖鑑.html")
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"Generated clean {out_html} ({os.path.getsize(out_html)//1024} KB)")
