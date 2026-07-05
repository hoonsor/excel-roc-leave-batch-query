import os

def convert_to_cp950():
    workspace_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
    src_dir = os.path.join(workspace_dir, "src")
    
    vba_files = ["clsLeaveRecord.cls", "modUtility.bas", "modImport.bas", "modQuery.bas"]
    
    for filename in vba_files:
        filepath = os.path.join(src_dir, filename)
        if not os.path.exists(filepath):
            print(f"File not found: {filepath}")
            continue
            
        print(f"Converting {filename} to CP950...")
        
        # 1. Read with UTF-8
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            print(f"  {filename} is already not UTF-8 or failed to decode. Reading with latin1 fallback...")
            with open(filepath, 'r', encoding='latin1') as f:
                content = f.read()
                
        # 2. Write back with CP950 (Taiwan Big5)
        # We replace unconvertible characters to prevent crash, though standard TC is fully supported
        with open(filepath, 'w', encoding='cp950', errors='replace') as f:
            f.write(content)
            
        print(f"  Successfully converted {filename} to CP950.")

if __name__ == "__main__":
    convert_to_cp950()
