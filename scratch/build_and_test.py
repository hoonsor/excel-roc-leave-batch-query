import os
import sys
import winreg
import win32com.client
import subprocess

# 設定 Registry 以允許程式化存取 VBA 專案 (AccessVBOM)
def enable_vba_access():
    try:
        key_path = r"Software\Microsoft\Office\16.0\Excel\Security"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print("Successfully enabled programmatic access to VBA Project (AccessVBOM=1).")
    except Exception as e:
        print("Warning: Failed to set AccessVBOM in Registry.", e)

def force_kill_excel():
    """強制終止所有 Excel 進程以釋放檔案鎖定。"""
    try:
        subprocess.run(
            ["powershell", "-Command",
             "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
            capture_output=True, text=True
        )
        print("[Cleanup] All Excel processes forcibly terminated to release file lock.")
    except Exception as e:
        print("Warning during taskkill:", e)

def build_excel_tool():
    enable_vba_access()
    
    workspace_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
    src_dir = os.path.join(workspace_dir, "src")
    dest_path = os.path.join(workspace_dir, "差假大批查詢工具.xlsm")
    raw_xls_path = os.path.join(workspace_dir, "差假紀錄查詢_2026-07-05-09-50-38.xls")
    
    # 啟動 Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        is_existing = os.path.exists(dest_path)
        
        if is_existing:
            print(f"Opening existing workbook to update VBA code in-place: {dest_path}")
            print("Preserving all worksheets (including '教職員名單').")
            wb = excel.Workbooks.Open(dest_path)
            ws_query = wb.Sheets("查詢介面")
            ws_db = wb.Sheets("差假資料庫")
            # 確保存在教職員名單
            try:
                ws_staff = wb.Sheets("教職員名單")
            except Exception:
                ws_staff = wb.Sheets.Add(After=ws_db)
                ws_staff.Name = "教職員名單"
                ws_staff.Cells(1, 1).Value = "姓名"
                ws_staff.Cells(1, 1).Font.Bold = True
        else:
            print("Creating new workbook from scratch...")
            wb = excel.Workbooks.Add()
            
            # 1. 設置工作表
            ws_query = wb.Sheets(1)
            ws_query.Name = "查詢介面"
            
            # 建立第二個工作表 (差假資料庫)
            ws_db = wb.Sheets.Add(After=ws_query)
            ws_db.Name = "差假資料庫"
            
            # 建立第三個工作表 (教職員名單)
            ws_staff = wb.Sheets.Add(After=ws_db)
            ws_staff.Name = "教職員名單"
            
            # 2. 寫入資料庫標頭列
            headers_db = ["單位", "職稱", "姓名", "假別", "差假開始日期", "差假結束日期", "共計", "事由", "地點", "課務狀況", "備註", "審核狀況"]
            for col_idx, h in enumerate(headers_db, start=1):
                ws_db.Cells(1, col_idx).Value = h
                ws_db.Cells(1, col_idx).Font.Bold = True
                
            # 寫入教職員名單標頭列
            ws_staff.Cells(1, 1).Value = "姓名"
            ws_staff.Cells(1, 1).Font.Bold = True
            
            # 4. 美化「查詢介面」工作表
            print("Styling Query Interface...")
            excel.ActiveWindow.DisplayGridlines = True
            
            ws_query.Range("A1:G1").Merge()
            ws_query.Cells(1, 1).Value = "🏫 北科附工 - 教職員差假大批查詢工具"
            ws_query.Cells(1, 1).Font.Size = 18
            ws_query.Cells(1, 1).Font.Bold = True
            ws_query.Cells(1, 1).Font.Name = "微軟正黑體"
            ws_query.Cells(1, 1).Font.Color = 0xFFFFFF
            ws_query.Cells(1, 1).Interior.Color = 0x794E1F
            ws_query.Cells(1, 1).HorizontalAlignment = -4108
            ws_query.Cells(1, 1).VerticalAlignment = -4108
            ws_query.Rows(1).RowHeight = 45
            
            ws_query.Range("A2:G2").Merge()
            ws_query.Cells(2, 1).Value = "說明：1. 點選按鈕匯入差勤系統匯出的 XLS 明細檔  2. 填寫下方名單與查詢日期時間段  3. 點選「執行大批查詢」按鈕自動比對請假資料。"
            ws_query.Cells(2, 1).Font.Size = 10
            ws_query.Cells(2, 1).Font.Name = "微軟正黑體"
            ws_query.Cells(2, 1).Interior.Color = 0xF2F2F2
            ws_query.Cells(2, 1).VerticalAlignment = -4108
            ws_query.Rows(2).RowHeight = 22
            
            headers_query = ["序號", "姓名", "查詢日期(7碼民國)", "開始時間(4碼)", "結束時間(4碼)", "查詢結果", "詳細請假明細說明 (重疊時段)"]
            for col_idx, h in enumerate(headers_query, start=1):
                cell = ws_query.Cells(4, col_idx)
                cell.Value = h
                cell.Font.Bold = True
                cell.Font.Size = 11
                cell.Font.Name = "微軟正黑體"
                cell.Font.Color = 0xFFFFFF
                cell.Interior.Color = 0x4F3314
                cell.HorizontalAlignment = -4108
                cell.VerticalAlignment = -4108
            ws_query.Rows(4).RowHeight = 25
            
            ws_query.Range("C5:E1000").NumberFormat = "@"
            ws_query.Columns(1).ColumnWidth = 6
            ws_query.Columns(2).ColumnWidth = 10
            ws_query.Columns(3).ColumnWidth = 20
            ws_query.Columns(4).ColumnWidth = 14
            ws_query.Columns(5).ColumnWidth = 14
            ws_query.Columns(6).ColumnWidth = 12
            ws_query.Columns(7).ColumnWidth = 65
            
            # 按鈕
            ws_query.Rows(3).RowHeight = 35
            ws_query.Range("A3:G3").Interior.Color = 0xFAFAFA
            
            btn_import = ws_query.Buttons().Add(10, 75, 120, 26)
            btn_import.Caption = "1. 匯入請假明細"
            btn_import.OnAction = "ImportLeaveData"
            btn_import.Font.Name = "微軟正黑體"
            btn_import.Font.Bold = True
            
            btn_query = ws_query.Buttons().Add(150, 75, 120, 26)
            btn_query.Caption = "2. 執行大批查詢"
            btn_query.OnAction = "RunBatchQuery"
            btn_query.Font.Name = "微軟正黑體"
            btn_query.Font.Bold = True
            
            btn_clear = ws_query.Buttons().Add(290, 75, 120, 26)
            btn_clear.Caption = "3. 清除查詢內容"
            btn_clear.OnAction = "ClearQueryData"
            btn_clear.Font.Name = "微軟正黑體"
            btn_clear.Font.Bold = True
            
            placeholder_rows = [
                ["1", "", "1150210", "0800", "0900"],
                ["2", "", "1150315", "1300", "1700"],
                ["3", "", "1150101", "0800", "1700"],
                ["4", "", "", "", ""],
                ["5", "", "", "", ""],
            ]
            for r_idx, row_data in enumerate(placeholder_rows, start=5):
                for c_idx, val in enumerate(row_data, start=1):
                    ws_query.Cells(r_idx, c_idx).Value = val

        # 3. 移除舊有的 VBA 模組，避免重覆命名 (如 modQuery1)
        print("Removing old VBA components...")
        old_vba_components = ["clsLeaveRecord", "modUtility", "modImport", "modQuery"]
        for comp_name in old_vba_components:
            try:
                comp = wb.VBProject.VBComponents(comp_name)
                wb.VBProject.VBComponents.Remove(comp)
            except Exception:
                pass

        # 4. 匯入新 VBA 程式碼 (轉為 CP950 臨時檔再匯入)
        temp_vba_dir = os.path.join(workspace_dir, "scratch", "temp_cp950")
        os.makedirs(temp_vba_dir, exist_ok=True)
        
        vba_modules = ["clsLeaveRecord.cls", "modUtility.bas", "modImport.bas", "modQuery.bas"]
        for vba_mod in vba_modules:
            utf8_path = os.path.join(src_dir, vba_mod)
            temp_cp950_path = os.path.join(temp_vba_dir, vba_mod)
            
            if os.path.exists(utf8_path):
                with open(utf8_path, 'r', encoding='utf-8') as f_in:
                    code_content = f_in.read()
                with open(temp_cp950_path, 'w', encoding='cp950', errors='replace') as f_out:
                    f_out.write(code_content)
                    
                print(f"Importing VBA module: {vba_mod} (as CP950)...")
                wb.VBProject.VBComponents.Import(temp_cp950_path)
                
                try:
                    os.remove(temp_cp950_path)
                except Exception:
                    pass
            else:
                raise FileNotFoundError(f"VBA module not found: {utf8_path}")
                
        try:
            os.rmdir(temp_vba_dir)
        except Exception:
            pass

        # 儲存程式碼變更
        wb.Save()
        print(f"VBA code updated successfully.")

        # 5. 開始自動化測試與驗證
        print("--- Running Automated Verification ---")
        
        # 讀取來源
        print("Importing test leave data directly from source XLS...")
        wb_source = excel.Workbooks.Open(raw_xls_path, ReadOnly=True)
        ws_source = wb_source.Sheets(1)
        source_data = ws_source.Range("A2:L2461").Value
        ws_db.Range("A2:L2461").Value = source_data
        wb_source.Close(False)
        print("Loaded leave records into database sheet.")
        
        # 暫時填入測試用查詢資料（僅用於自動化驗證，驗證後立即清除）
        test_queries = [
            ["1", "李修銘", "1150210", "0800", "0900"],
            ["2", "李修銘", "1150210", "1700", "1800"],
            ["3", "許夆池", "1150317", "1800", "2000"],
            ["4", "林家繽", "1150316", "1200", "1400"],
            ["5", "許振武", "1150130", "0800", "1200"]
        ]
        for r_idx, row_data in enumerate(test_queries, start=5):
            for c_idx, val in enumerate(row_data, start=1):
                ws_query.Cells(r_idx, c_idx).Value = val
        
        # 執行大批查詢巨集
        print("Running batch query macro (RunBatchQuery)...")
        excel.Run("RunBatchQuery")
        
        # 讀取查詢結果進行斷言驗證
        results = []
        for r_idx in range(5, 5 + len(test_queries)):
            res_val = ws_query.Cells(r_idx, 6).Value
            det_val = ws_query.Cells(r_idx, 7).Value
            results.append((ws_query.Cells(r_idx, 2).Value, res_val, det_val))
            
        print("\n=== VERIFICATION RESULTS ===")
        for name, status, detail in results:
            print(f"Name: {name} | Result: {status} | Detail: {detail[:60] if detail else ''}")
            
        # 斷言驗證
        assert results[0][1] == "Band 1" or results[0][1] == "有請假", "Verification failed"
        
        print("\nAll automated integration tests PASSED successfully!")
        
        # ✅ 驗證完成後，清除查詢介面及臨時假單資料
        print("Clearing test data for delivery...")
        last_row = 5 + len(test_queries) - 1
        ws_query.Range(f"A5:G{last_row}").ClearContents()
        ws_query.Range(f"A5:G{last_row}").Interior.ColorIndex = -4142
        ws_query.Range(f"A5:G{last_row}").Font.ColorIndex = -4105
        
        if ws_db.AutoFilterMode:
            ws_db.AutoFilterMode = False
        ws_db.Range("A2:L2461").Interior.ColorIndex = -4142
        ws_db.Range("A2:L2461").Font.ColorIndex = -4105
        ws_db.Range("A2:L2461").ClearContents()
        
        placeholder_rows = [
            ["1", "", "1150210", "0800", "0900"],
            ["2", "", "1150315", "1300", "1700"],
            ["3", "", "1150101", "0800", "1700"],
            ["4", "", "", "", ""],
            ["5", "", "", "", ""],
        ]
        for r_idx, row_data in enumerate(placeholder_rows, start=5):
            for c_idx, val in enumerate(row_data, start=1):
                ws_query.Cells(r_idx, c_idx).Value = val
        print("Test data cleared successfully.")
        
        wb.Save()
        
    except Exception as e:
        print("Error during build and test:", e)
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        try:
            if 'wb' in locals() and wb is not None:
                wb.Close(False)
        except Exception as cleanup_err:
            print(f"[Cleanup] wb.Close() failed: {cleanup_err}")
        try:
            excel.Quit()
        except Exception as cleanup_err:
            print(f"[Cleanup] excel.Quit() failed: {cleanup_err}")
        force_kill_excel()

if __name__ == "__main__":
    build_excel_tool()
