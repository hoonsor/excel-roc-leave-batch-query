Attribute VB_Name = "modImport"
Option Explicit

''' <summary>
''' 匯入外部差假明細檔案
''' </summary>
Public Sub ImportLeaveData()
    Dim fd As Object
    Dim fileSelected As String
    
    ' 1. 彈出檔案選取對話框
    fileSelected = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "請選取差勤系統匯出的請假明細檔")
    
    If fileSelected = "False" Then
        MsgBox "已取消匯入操作。", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 2. 效能優化開始
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    ' 3. 設定目標工作表 (在本活頁簿中)
    Set wsDest = ThisWorkbook.Sheets("差假資料庫")
    
    ' 清除目標工作表的舊資料 (保留第一行標頭)
    Dim lastRowDest As Long
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDest > 1 Then
        wsDest.Range("A2:L" & lastRowDest).ClearContents
    End If
    
    ' 4. 開啟來源活頁簿 (唯讀)
    Set wbSource = Workbooks.Open(Filename:=fileSelected, ReadOnly:=True)
    
    ' 防呆：尋找適當的工作表（通常是第一個，或名稱含有 "差假" 或 "Sheet"）
    Dim sheetFound As Boolean
    Dim wsLoop As Worksheet
    sheetFound = False
    
    For Each wsLoop In wbSource.Worksheets
        If InStr(wsLoop.Name, "差假") > 0 Or InStr(wsLoop.Name, "Sheet") > 0 Then
            Set wsSource = wsLoop
            sheetFound = True
            Exit For
        End If
    Next wsLoop
    
    ' 若沒找到特定名稱，取第一個工作表
    If Not sheetFound Then
        Set wsSource = wbSource.Worksheets(1)
    End If
    
    ' 5. 掃描與匯入資料
    Dim sourceLastRow As Long
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' 判斷標頭行位置
    Dim headerRow As Long
    Dim r As Long
    headerRow = 1 ' 預設為第 1 行
    
    ' 掃描前 10 行尋找包含 "單位" 和 "姓名" 的標頭列
    For r = 1 To MinValue(10, sourceLastRow)
        If InStr(wsSource.Cells(r, 1).Value, "單位") > 0 And (InStr(wsSource.Cells(r, 3).Value, "姓名") > 0 Or InStr(wsSource.Cells(r, 2).Value, "職稱") > 0) Then
            headerRow = r
            Exit For
        End If
    Next r
    
    Dim destRow As Long
    destRow = 2 ' 從第 2 行開始寫入
    
    ' 從標頭列下一行開始讀取
    Dim colCount As Integer
    colCount = 12 ' 共有 12 個欄位
    
    Dim cellVal As String
    For r = (headerRow + 1) To sourceLastRow
        cellVal = Trim(CStr(wsSource.Cells(r, 1).Value))
        
        ' 排除統計彙總行 (開頭為 ■ 或包含 '共計'/'■')
        If Left(cellVal, 1) = "■" Or InStr(cellVal, "■") > 0 Then
            ' 跳過統計列
            GoTo NextRow
        End If
        
        ' 排除空行 (如果單位、姓名與日期皆為空)
        If cellVal = "" And Trim(CStr(wsSource.Cells(r, 3).Value)) = "" And Trim(CStr(wsSource.Cells(r, 5).Value)) = "" Then
            GoTo NextRow
        End If
        
        ' 複製該行資料 (1 到 12 欄)
        Dim c As Integer
        For c = 1 To colCount
            wsDest.Cells(destRow, c).Value = wsSource.Cells(r, c).Value
        Next c
        
        destRow = destRow + 1
        
NextRow:
    Next r
    
    ' 6. 關閉來源活頁簿
    wbSource.Close SaveChanges:=False
    
    ' 7. 效能優化結束與恢復
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "差假明細匯入完成！共匯入 " & (destRow - 2) & " 筆有效資料。", vbInformation, "匯入成功"
    Exit Sub

ErrorHandler:
    ' 發生錯誤時確保恢復 Excel 環境
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "匯入請假明細時發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function MinValue(val1 As Long, val2 As Long) As Long
    If val1 < val2 Then
        MinValue = val1
    Else
        MinValue = val2
    End If
End Function
