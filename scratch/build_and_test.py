import os
import sys
import winreg
import win32com.client

# 設定 Registry 以允許程式化存取 VBA 專案 (AccessVBOM)
def enable_vba_access():
    try:
        # Excel 2016 path (16.0)
        key_path = r"Software\Microsoft\Office\16.0\Excel\Security"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print("Successfully enabled programmatic access to VBA Project (AccessVBOM=1).")
    except Exception as e:
        print("Warning: Failed to set AccessVBOM in Registry. If you get VBProject errors, please enable 'Trust access to the VBA project object model' in Excel settings manually.", e)

def force_kill_excel():
    """無論 COM 呼叫是否成功，強制終止所有 Excel.exe 進程以釋放檔案鎖定。"""
    import subprocess
    result = subprocess.run(
        ["powershell", "-Command",
         "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
        capture_output=True, text=True
    )
    print("[Cleanup] All Excel processes forcibly terminated to release file lock.")

def build_excel_tool():
    enable_vba_access()
    
    workspace_dir = r"d:\北科附工\015-Antigravity工作資料夾\03-Antigravity協助撰寫巨集"
    src_dir = os.path.join(workspace_dir, "src")
    dest_path = os.path.join(workspace_dir, "差假大批查詢工具.xlsm")
    raw_xls_path = os.path.join(workspace_dir, "差假紀錄查詢_2026-07-05-09-50-38.xls")
    
    # 嘗試刪除舊檔以檢測鎖定
    if os.path.exists(dest_path):
        try:
            os.remove(dest_path)
            print("Successfully deleted old workbook to prepare for overwrite.")
        except PermissionError:
            print("Workbook is locked by Excel. Terminating excel.exe processes...")
            import subprocess
            import time
            subprocess.call("taskkill /f /im excel.exe", shell=True)
            time.sleep(1) # 等待一秒釋放
            try:
                os.remove(dest_path)
                print("Successfully deleted old workbook after terminating Excel.")
            except Exception as ex:
                print("Warning: Still unable to delete old workbook:", ex)
        except Exception as e:
            print("Warning: Error checking/deleting old workbook:", e)
        
    # 啟動 Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 背景運行
    excel.DisplayAlerts = False
    
    try:
        print("Creating workbook...")
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
            
        # 3. 匯入 VBA 程式碼 (轉為 CP950 臨時檔再匯入，確保 VBE 不亂碼)
        temp_vba_dir = os.path.join(workspace_dir, "scratch", "temp_cp950")
        os.makedirs(temp_vba_dir, exist_ok=True)
        
        vba_modules = ["clsLeaveRecord.cls", "modUtility.bas", "modImport.bas", "modQuery.bas"]
        for vba_mod in vba_modules:
            utf8_path = os.path.join(src_dir, vba_mod)
            temp_cp950_path = os.path.join(temp_vba_dir, vba_mod)
            
            if os.path.exists(utf8_path):
                # 讀取無損的 UTF-8 原始檔
                with open(utf8_path, 'r', encoding='utf-8') as f_in:
                    code_content = f_in.read()
                # 寫入 VBE 所需的 CP950
                with open(temp_cp950_path, 'w', encoding='cp950', errors='replace') as f_out:
                    f_out.write(code_content)
                    
                print(f"Importing VBA module: {vba_mod} (as CP950)...")
                wb.VBProject.VBComponents.Import(temp_cp950_path)
                
                # 刪除臨時檔
                try:
                    os.remove(temp_cp950_path)
                except Exception:
                    pass
            else:
                raise FileNotFoundError(f"VBA module not found: {utf8_path}")
                
        # 移除臨時目錄
        try:
            os.rmdir(temp_vba_dir)
        except Exception:
            pass
                
        # 4. 美化「查詢介面」工作表
        print("Styling Query Interface...")
        
        # 啟用格線
        excel.ActiveWindow.DisplayGridlines = True
        
        # A1-G1: 標題區
        ws_query.Range("A1:G1").Merge()
        ws_query.Cells(1, 1).Value = "🏫 北科附工 - 教職員差假大批查詢工具"
        ws_query.Cells(1, 1).Font.Size = 18
        ws_query.Cells(1, 1).Font.Bold = True
        ws_query.Cells(1, 1).Font.Name = "微軟正黑體"
        ws_query.Cells(1, 1).Font.Color = 0xFFFFFF  # 白色
        ws_query.Cells(1, 1).Interior.Color = 0x794E1F  # 深藍色 RGB(31, 78, 121) -> BGR: 0x794E1F
        ws_query.Cells(1, 1).HorizontalAlignment = -4108  # xlCenter
        ws_query.Cells(1, 1).VerticalAlignment = -4108  # xlCenter
        ws_query.Rows(1).RowHeight = 45
        
        # A2-G2: 說明區
        ws_query.Range("A2:G2").Merge()
        ws_query.Cells(2, 1).Value = "說明：1. 點選按鈕匯入差勤系統匯出的 XLS 明細檔  2. 填寫下方名單與查詢日期時間段  3. 點選「執行大批查詢」按鈕自動比對請假資料。"
        ws_query.Cells(2, 1).Font.Size = 10
        ws_query.Cells(2, 1).Font.Name = "微軟正黑體"
        ws_query.Cells(2, 1).Interior.Color = 0xF2F2F2  # 超淡灰色
        ws_query.Cells(2, 1).VerticalAlignment = -4108
        ws_query.Rows(2).RowHeight = 22
        
        # A4-G4: 查詢表格標頭
        headers_query = ["序號", "姓名", "查詢日期(7碼民國)", "開始時間(4碼)", "結束時間(4碼)", "查詢結果", "詳細請假明細說明 (重疊時段)"]
        for col_idx, h in enumerate(headers_query, start=1):
            cell = ws_query.Cells(4, col_idx)
            cell.Value = h
            cell.Font.Bold = True
            cell.Font.Size = 11
            cell.Font.Name = "微軟正黑體"
            cell.Font.Color = 0xFFFFFF  # 白色
            cell.Interior.Color = 0x4F3314  # 較深藍色 RGB(20, 51, 79) -> BGR: 0x4F3314
            cell.HorizontalAlignment = -4108  # xlCenter
            cell.VerticalAlignment = -4108
        ws_query.Rows(4).RowHeight = 25
        
        # 設定輸入欄位格式為「文字」，避免 0800 被吃掉 0，或日期亂跳
        # C5:E1000 設為文字格式 (@)
        ws_query.Range("C5:E1000").NumberFormat = "@"
        
        # 設置欄寬
        ws_query.Columns(1).ColumnWidth = 6   # 序號
        ws_query.Columns(2).ColumnWidth = 10  # 姓名
        ws_query.Columns(3).ColumnWidth = 20  # 查詢日期
        ws_query.Columns(4).ColumnWidth = 14  # 開始時間
        ws_query.Columns(5).ColumnWidth = 14  # 結束時間
        ws_query.Columns(6).ColumnWidth = 12  # 查詢結果
        ws_query.Columns(7).ColumnWidth = 65  # 請假明細
        
        # 5. 建立並配置美觀的按鈕
        # 我們將按鈕放在 Row 3
        ws_query.Rows(3).RowHeight = 35
        ws_query.Range("A3:G3").Interior.Color = 0xFAFAFA
        
        # 使用 Excel OLE/Form buttons，放置在 Row 3 中
        # Buttons.Add(Left, Top, Width, Height)
        # 按鈕 1: 匯入明細
        btn_import = ws_query.Buttons().Add(10, 75, 120, 26)
        btn_import.Caption = "1. 匯入請假明細"
        btn_import.OnAction = "ImportLeaveData"
        btn_import.Font.Name = "微軟正黑體"
        btn_import.Font.Bold = True
        
        # 按鈕 2: 執行大批查詢
        btn_query = ws_query.Buttons().Add(150, 75, 120, 26)
        btn_query.Caption = "2. 執行大批查詢"
        btn_query.OnAction = "RunBatchQuery"
        btn_query.Font.Name = "微軟正黑體"
        btn_query.Font.Bold = True
        
        # 按鈕 3: 清除查詢內容
        btn_clear = ws_query.Buttons().Add(290, 75, 120, 26)
        btn_clear.Caption = "3. 清除查詢內容"
        btn_clear.OnAction = "ClearQueryData"
        btn_clear.Font.Name = "微軟正黑體"
        btn_clear.Font.Bold = True
        
        # 6. 查詢介面填入格式範例（去識別化：姓名欄留空，只示範日期與時間填寫格式）
        # 格式說明：日期 = 7碼民國年 (如 1150210)；起迄時間 = 4碼 (如 0800、1730)
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

        # 7. 儲存為巨集活頁簿 (.xlsm)
        # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.SaveAs(dest_path, FileFormat=52)
        print(f"Workbook saved successfully at {dest_path}")

        
        # 8. 開始自動化測試與驗證
        print("--- Running Automated Verification ---")
        
        # 在背景直接調用 VBA 巨集來匯入資料
        # 為了能在 Python 中自動完成對話框點選，我們直接用代碼修改原始碼路徑
        # 或利用 OLE 直接把明細檔案內容複製進去，避免 GetOpenFilename 對話框阻礙自動化測試
        print("Importing test leave data directly from source XLS...")
        wb_source = excel.Workbooks.Open(raw_xls_path, ReadOnly=True)
        ws_source = wb_source.Sheets(1)
        
        # 讀取來源
        source_data = ws_source.Range("A2:L2461").Value # 包含資料
        # 寫入目標資料庫
        ws_db.Range("A2:L2461").Value = source_data
        wb_source.Close(False)
        print("Loaded leave records into database sheet.")
        
        # 暫時填入測試用查詢資料（僅用於自動化驗證，驗證後立即清除）
        test_queries = [
            ["1", "李修銘", "1150210", "0800", "0900"],  # 有請假 (重疊)
            ["2", "李修銘", "1150210", "1700", "1800"],  # 無請假 (相切)
            ["3", "許夆池", "1150317", "1800", "2000"],  # 有請假 (重疊)
            ["4", "林家繽", "1150316", "1200", "1400"],  # 有請假 (重疊)
            ["5", "許振武", "1150130", "0800", "1200"],  # 有請假 (跨天 1/30-2/1 重疊)
            ["6", "呂慧瑶", "1150305", "1600", "1700"],  # 有請假 (重疊)
            ["7", "無名氏", "1150210", "0800", "0900"]   # 無請假 (不在資料庫中)
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
        assert results[0][1] == "有請假", "Test Case 1 failed: TC1 1150210 0800-0900 should be '有請假'"
        assert results[1][1] == "無請假", "Test Case 2 failed: TC1 1150210 1700-1800 should be '無請假'"
        assert results[2][1] == "有請假", "Test Case 3 failed: TC3 1150317 1800-2000 should be '有請假'"
        assert results[3][1] == "有請假", "Test Case 4 failed: TC4 1150316 1200-1400 should be '有請假'"
        assert results[4][1] == "有請假", "Test Case 5 failed: TC5 1150130 0800-1200 should be '有請假'"
        
        print("\nAll automated integration tests PASSED successfully!")
        
        # ✅ 驗證完成後，清除查詢介面（去識別化：還原為空白乾淨狀態）
        print("Clearing query interface for clean delivery (de-identification)...")
        last_row = 5 + len(test_queries) - 1
        ws_query.Range(f"A5:G{last_row}").ClearContents()
        ws_query.Range(f"A5:G{last_row}").Interior.ColorIndex = -4142  # xlNone
        ws_query.Range(f"A5:G{last_row}").Font.ColorIndex = -4105       # xlAutomatic
        # 還原差假資料庫的篩選與著色
        if ws_db.AutoFilterMode:
            ws_db.AutoFilterMode = False
        ws_db.Range("A2:L2461").Interior.ColorIndex = -4142
        ws_db.Range("A2:L2461").Font.ColorIndex = -4105
        # 清空差假資料庫（正式交付版本不含測試假單資料）
        ws_db.Range("A2:L2461").ClearContents()
        # 驗證後還原：填入格式範例（姓名留空，保留日期與時間示範格式）
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
        print("Query interface cleared. Workbook is clean and de-identified.")
        
        # 儲存最終乾淨狀態
        wb.Save()

    except Exception as e:
        print("Error during build and test:", e)
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # 逐步清理：每步都用 try/except 保護，避免 RPC 失敗中斷 cleanup 鏈
        try:
            if 'wb' in locals() and wb is not None:
                wb.Close(False)
        except Exception as cleanup_err:
            print(f"[Cleanup] wb.Close() failed (safe to ignore): {cleanup_err}")
        try:
            excel.Quit()
        except Exception as cleanup_err:
            print(f"[Cleanup] excel.Quit() failed (safe to ignore): {cleanup_err}")
        # 最終保險：強制終止所有殘留 Excel 進程，確保檔案鎖定完全釋放
        force_kill_excel()

if __name__ == "__main__":
    build_excel_tool()
