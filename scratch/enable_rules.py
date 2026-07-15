# -*- coding: utf-8 -*-
import os

base_dir = r"D:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
rules_file = os.path.join(base_dir, ".agents", "rules", "roc_pdf_rules.py")

if not os.path.exists(rules_file):
    print(f"Error: Rules file {rules_file} does not exist.")
    exit(1)

with open(rules_file, "r", encoding="utf-8") as f:
    content = f.read()

# Replace IS_ENABLED = False with IS_ENABLED = True
if "IS_ENABLED = False" in content:
    content = content.replace("IS_ENABLED = False", "IS_ENABLED = True")
    with open(rules_file, "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully enabled ROC PDF rules (IS_ENABLED = True). Format learning completed!")
else:
    print("Rules are already enabled or the target line was not found.")
