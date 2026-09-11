import sys
import pprint

sys.stdout.reconfigure(encoding='utf-8')

# Read records from records_out.py
with open('records_out.py', 'r', encoding='utf-8') as f:
    code = f.read()
    
# Execute the code to get the records list
local_vars = {}
exec(code, {}, local_vars)
records = local_vars['records']

# Format to string
records_str = ',\n'.join(['        ' + repr(r) for r in records])

with open('.agents/rules/roc_pdf_rules.py', 'r', encoding='utf-8') as f:
    content = f.read()

insert_idx = content.rfind('}')
if insert_idx != -1:
    new_content = content[:insert_idx] + '    # 代理導師職務印領名冊\n    "代理導師職務印領名冊": [\n' + records_str + '\n    ],\n' + content[insert_idx:]
    with open('.agents/rules/roc_pdf_rules.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('Successfully added to roc_pdf_rules.py')
else:
    print('Failed to find insertion point')
