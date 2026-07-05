Attribute VB_Name = "modQuery"
Option Explicit

' 定義查詢介面的欄位常數
Private Const COL_SEQ As Integer = 1      ' 欄位 A: 序號
Private Const COL_NAME As Integer = 2     ' 欄位 B: 姓名
Private Const COL_DATE As Integer = 3     ' 欄位 C: 查詢日期
Private Const COL_TSTART As Integer = 4   ' 欄位 D: 開始時間
Private Const COL_TEND As Integer = 5     ' 欄位 E: 結束時間
Private Const COL_RESULT As Integer = 6   ' 欄位 F: 查詢結果
Private Const COL_DETAIL As Integer = 7   ' 欄位 G: 請假明細

''' <summary>
''' 執行大批請假查詢
''' </summary>
Public Sub RunBatchQuery()
    Dim wsQuery As Worksheet
    Dim wsDb As Worksheet
    
    Set wsQuery = ThisWorkbook.Sheets("查詢介面")
    Set wsDb = ThisWorkbook.Sheets("差假資料庫")
    
    ' 1. 取得查詢列數
    Dim lastRowQuery As Long
    lastRowQuery = wsQuery.Cells(wsQuery.Rows.Count, COL_NAME).End(xlUp).Row
    
    If lastRowQuery < 5 Then
        MsgBox "未偵測到待查詢的教職員姓名。請從第 5 行開始填寫資料。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 2. 效能優化開始
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 3. 將差假資料庫載入到記憶體中並建立字典索引 (以姓名為 Key)
    Dim dictDb As Object
    Set dictDb = CreateObject("Scripting.Dictionary")
    
    Dim lastRowDb As Long
    lastRowDb = wsDb.Cells(wsDb.Rows.Count, 1).End(xlUp).Row
    
    If lastRowDb >= 2 Then
        Dim arrDb() As Variant
        arrDb = wsDb.Range("A2:L" & lastRowDb).Value
        
        Dim rDb As Long
        Dim empName As String
        Dim rec As clsLeaveRecord
        Dim colLeaves As Collection
        
        For rDb = 1 To UBound(arrDb, 1)
            empName = Trim(CStr(arrDb(rDb, 3))) ' 姓名 (第 3 欄)
            If empName <> "" Then
                ' 只載入「已簽核」或「審核完成」的假單，過濾未簽核的
                Dim strStatus As String
                strStatus = Trim(CStr(arrDb(rDb, 12))) ' 審核狀況 (第 12 欄)
                
                If strStatus = "已簽核" Or strStatus = "審核完成" Or strStatus = "" Then
                    ' 實例化類別物件
                    Set rec = New clsLeaveRecord
                    rec.Status = strStatus
                    rec.LeaveType = CStr(arrDb(rDb, 4)) ' 假別 (第 4 欄)
                    rec.RawStart = CStr(arrDb(rDb, 5))  ' 開始日期 (第 5 欄)
                    rec.RawEnd = CStr(arrDb(rDb, 6))    ' 結束日期 (第 6 欄)
                    rec.Reason = CStr(arrDb(rDb, 8))    ' 事由 (第 8 欄)
                    
                    ' 解析日期時間為西元 Date
                    On Error Resume Next
                    rec.StartDT = modUtility.ParseRocDateTime(rec.RawStart)
                    rec.EndDT = modUtility.ParseRocDateTime(rec.RawEnd)
                    On Error GoTo ErrorHandler
                    
                    ' 放入字典
                    If Not dictDb.Exists(empName) Then
                        Set colLeaves = New Collection
                        dictDb.Add empName, colLeaves
                    Else
                        Set colLeaves = dictDb(empName)
                    End If
                    colLeaves.Add rec
                End If
            End If
        Next rDb
    End If
    
    ' 4. 清除先前的查詢結果與格式
    wsQuery.Range("F5:G" & lastRowQuery).ClearContents
    wsQuery.Range("A5:G" & lastRowQuery).Interior.ColorIndex = xlNone
    wsQuery.Range("A5:G" & lastRowQuery).Font.ColorIndex = xlAutomatic
    
    ' 5. 開始逐列查詢比對
    Dim rQuery As Long
    Dim qName As String, qDateStr As String, qStartStr As String, qEndStr As String
    Dim qStartDT As Date, qEndDT As Date
    Dim queryValid As Boolean
    Dim errMsg As String
    
    For rQuery = 5 To lastRowQuery
        queryValid = True
        errMsg = ""
        
        qName = Trim(CStr(wsQuery.Cells(rQuery, COL_NAME).Value))
        qDateStr = Trim(CStr(wsQuery.Cells(rQuery, COL_DATE).Value))
        qStartStr = Trim(CStr(wsQuery.Cells(rQuery, COL_TSTART).Value))
        qEndStr = Trim(CStr(wsQuery.Cells(rQuery, COL_TEND).Value))
        
        ' 若該列姓名為空，跳過
        If qName = "" Then GoTo NextQuery
        
        ' 欄位完整度檢查
        If qDateStr = "" Or qStartStr = "" Or qEndStr = "" Then
            queryValid = False
            errMsg = "輸入資料不完整"
            GoTo WriteResult
        End If
        
        ' 驗證時間格式並補足長度 (4碼數字，例如 800 -> 0800)
        If Len(qStartStr) < 4 And IsNumeric(qStartStr) Then qStartStr = Format(CInt(qStartStr), "0000")
        If Len(qEndStr) < 4 And IsNumeric(qEndStr) Then qEndStr = Format(CInt(qEndStr), "0000")
        
        If Len(qStartStr) <> 4 Or Not IsNumeric(qStartStr) Or Len(qEndStr) <> 4 Or Not IsNumeric(qEndStr) Then
            queryValid = False
            errMsg = "時間格式錯誤 (須為4碼)"
            GoTo WriteResult
        End If
        
        ' 解析查詢起點與訖點為西元 Date
        Dim dtBase As Date
        On Error Resume Next
        dtBase = modUtility.CDateFromROC(qDateStr)
        If Err.Number <> 0 Then
            queryValid = False
            errMsg = "日期格式錯誤 (須7碼)"
            Err.Clear
            On Error GoTo ErrorHandler
            GoTo WriteResult
        End If
        
        Dim startHour As Integer, startMin As Integer
        Dim endHour As Integer, endMin As Integer
        
        startHour = CInt(Left(qStartStr, 2))
        startMin = CInt(Right(qStartStr, 2))
        endHour = CInt(Left(qEndStr, 2))
        endMin = CInt(Right(qEndStr, 2))
        
        If startHour < 0 Or startHour > 23 Or startMin < 0 Or startMin > 59 Or _
           endHour < 0 Or endHour > 23 Or endMin < 0 Or endMin > 59 Then
            queryValid = False
            errMsg = "時間範圍不合法 (0000-2359)"
            On Error GoTo ErrorHandler
            GoTo WriteResult
        End If
        
        qStartDT = DateAdd("h", startHour, dtBase)
        qStartDT = DateAdd("n", startMin, qStartDT)
        
        qEndDT = DateAdd("h", endHour, dtBase)
        qEndDT = DateAdd("n", endMin, qEndDT)
        
        If qEndDT <= qStartDT Then
            queryValid = False
            errMsg = "迄點時間須大於起點"
            On Error GoTo ErrorHandler
            GoTo WriteResult
        End If
        On Error GoTo ErrorHandler
        
WriteResult:
        If Not queryValid Then
            wsQuery.Cells(rQuery, COL_RESULT).Value = errMsg
            wsQuery.Range(wsQuery.Cells(rQuery, 1), wsQuery.Cells(rQuery, COL_DETAIL)).Interior.Color = RGB(255, 242, 204) ' 橙黃色警告
            GoTo NextQuery
        End If
        
        ' 6. 核心比對邏輯 (使用 Dictionary + Collection 進行比對)
        Dim hasLeave As Boolean
        Dim leaveDetails As String
        hasLeave = False
        leaveDetails = ""
        
        If dictDb.Exists(qName) Then
            Set colLeaves = dictDb(qName)
            Dim item As Variant
            Dim lRec As clsLeaveRecord
            
            For Each item In colLeaves
                Set lRec = item
                
                ' 重疊演算法：[S_query, E_query] 與 [S_leave, E_leave] 重疊
                ' 條件：S_query < E_leave 且 E_query > S_leave
                If qStartDT < lRec.EndDT And qEndDT > lRec.StartDT Then
                    hasLeave = True
                    ' 組合顯示資訊： 假別 (開始時間 ~ 結束時間) 事由
                    Dim strPeriod As String
                    strPeriod = lRec.LeaveType & " (" & lRec.RawStart & " ~ " & lRec.RawEnd & ")"
                    If Trim(lRec.Reason) <> "" Then
                        strPeriod = strPeriod & " 事由: " & Trim(lRec.Reason)
                    End If
                    
                    If leaveDetails = "" Then
                        leaveDetails = strPeriod
                    Else
                        leaveDetails = leaveDetails & vbCrLf & strPeriod
                    End If
                End If
            Next item
        End If
        
        ' 寫回結果並美化著色
        If hasLeave Then
            wsQuery.Cells(rQuery, COL_RESULT).Value = "有請假"
            wsQuery.Cells(rQuery, COL_DETAIL).Value = leaveDetails
            ' 淡紅色背景，紅色文字 (著色整列查詢資料)
            wsQuery.Range(wsQuery.Cells(rQuery, 1), wsQuery.Cells(rQuery, COL_DETAIL)).Interior.Color = RGB(253, 233, 230)
            wsQuery.Range(wsQuery.Cells(rQuery, 1), wsQuery.Cells(rQuery, COL_DETAIL)).Font.Color = RGB(220, 20, 60)
        Else
            wsQuery.Cells(rQuery, COL_RESULT).Value = "無請假"
            ' 綠色調美化
            wsQuery.Range(wsQuery.Cells(rQuery, COL_RESULT), wsQuery.Cells(rQuery, COL_RESULT)).Font.Color = RGB(34, 139, 34)
        End If
        
NextQuery:
    Next rQuery
    
    ' 7. 效能優化結束與恢復
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "大批查詢比對完成！", vbInformation, "查詢成功"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "執行大批查詢時發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

''' <summary>
''' 清除查詢介面的輸入與結果
''' </summary>
Public Sub ClearQueryData()
    Dim wsQuery As Worksheet
    Set wsQuery = ThisWorkbook.Sheets("查詢介面")
    
    Dim lastRowQuery As Long
    lastRowQuery = wsQuery.Cells(wsQuery.Rows.Count, COL_NAME).End(xlUp).Row
    
    If lastRowQuery >= 5 Then
        ' 彈出確認對話框
        Dim ans As VbMsgBoxResult
        ans = MsgBox("您確定要清除所有的查詢輸入與結果嗎？", vbQuestion + vbYesNo, "確認清除")
        If ans = vbNo Then Exit Sub
        
        Application.ScreenUpdating = False
        ' 清除 B5 到 G最後一列
        wsQuery.Range("B5:G" & lastRowQuery).ClearContents
        wsQuery.Range("A5:G" & lastRowQuery).Interior.ColorIndex = xlNone
        wsQuery.Range("A5:G" & lastRowQuery).Font.ColorIndex = xlAutomatic
        Application.ScreenUpdating = True
    End If
End Sub
