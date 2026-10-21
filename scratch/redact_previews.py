import os
from PIL import Image, ImageDraw
import pypdfium2 as pdfium

base_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
out_dir = os.path.join(base_dir, "docs", "style_previews")
os.makedirs(out_dir, exist_ok=True)

samples = [
    ("style_01_patrol.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "1150708-職場巡輔交通費申請表-廖建清.pdf")),
    ("style_02_level_c.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "1150918-115年丙級專案技能檢定清冊（1月）.pdf")),
    ("style_03_special_ed.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "●使用過之測試資料", "1150708-身心障礙學習扶助印領清冊鐘點費-1.pdf")),
    ("style_04_industry.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "●使用過之測試資料", "1150708-產業特殊需求類科教學鐘點費.pdf")),
    ("style_05_english.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "1150715-英文科鐘點費-1.pdf")),
    ("style_06_club.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "1150909-115-1-8月聯課活動鐘點費印領清冊.pdf")),
    ("style_07_homeroom.png", os.path.join(base_dir, "001-歷史檔案（已作業完成）", "1150911-115年8月代理導師職務印領名冊(0817~0831).pdf")),
    ("style_08_national_test.png", os.path.join(base_dir, "003-測試資料", "1151019-115年度第2梯次全國檢定建築物室內設計乙級術科技術士檢定術科（檢定當日與非檢定當日工作人員）人員酬勞費印領清冊.pdf")),
]

# Normalized bounding boxes (x1, y1, x2, y2) in [0.0, 1.0] to fill with solid white (#FFFFFF)
redactions = {
    "style_01_patrol.png": [
        (0.150, 0.250, 0.194, 0.395), # 姓名欄: 廖建清
        (0.675, 0.736, 0.880, 0.775), # 簽名欄: 廖建清
        (0.215, 0.785, 0.365, 0.835), # 申請人: 廖建清
        (0.235, 0.840, 0.388, 0.888), # 特教組長章
        (0.235, 0.900, 0.388, 0.965), # 教務主任章
    ],
    "style_02_level_c.png": [
        (0.131, 0.262, 0.188, 0.953), # 姓名欄 (編號1~13)
        (0.780, 0.968, 0.890, 0.999), # 右下角職章
    ],
    "style_03_special_ed.png": [
        (0.140, 0.172, 0.222, 0.285), # 授課教師欄: 吳珮雯
        (0.148, 0.338, 0.295, 0.448), # 承辦人/組長/主任章
        (0.308, 0.355, 0.460, 0.455), # 人事室/人事主任章
    ],
    "style_04_industry.png": [
        (0.100, 0.130, 0.218, 0.940), # 授課教師欄 (序號1~11)
    ],
    "style_05_english.png": [
        (0.116, 0.263, 0.188, 0.410), # 姓名欄 (項次1~3)
        (0.080, 0.763, 0.215, 0.915), # 製表/主任章
        (0.422, 0.790, 0.528, 0.840), # 出納組章
    ],
    "style_06_club.png": [
        (0.268, 0.220, 0.320, 0.885), # 指導老師欄姓名 (保留「老師」二字或整欄遮蔽)
    ],
    "style_07_homeroom.png": [
        (0.112, 0.160, 0.163, 0.388), # 代理導師姓名欄 (編號1~9)
        (0.283, 0.160, 0.335, 0.388), # 原導師姓名欄 (編號1~9)
    ],
    "style_08_national_test.png": [
        (0.062, 0.248, 0.140, 0.822), # 姓名欄 (全16列)
    ],
}

for out_name, pdf_path in samples:
    doc = pdfium.PdfDocument(pdf_path)
    page = doc[0]
    bitmap = page.render(scale=2)
    pil_img = bitmap.to_pil().convert("RGB")
    doc.close()
    
    w, h = pil_img.size
    draw = ImageDraw.Draw(pil_img)
    for (x1, y1, x2, y2) in redactions.get(out_name, []):
        box = [int(x1 * w), int(y1 * h), int(x2 * w), int(y2 * h)]
        draw.rectangle(box, fill="#FFFFFF")
        
    out_path = os.path.join(out_dir, out_name)
    pil_img.save(out_path, optimize=True)
    print(f"Saved redacted {out_name}")
