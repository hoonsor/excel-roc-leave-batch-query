Attribute VB_Name = "modUtility"
Option Explicit

' 台灣中華民國曆與西元偏差值常數
Public Const ROC_YEAR_OFFSET As Integer = 1911

''' <summary>
''' 判斷民國年是否為閏年
''' </summary>
Public Function IsRocLeapYear(ByVal intRocYear As Integer) As Boolean
    Dim intAdYear As Integer
    intAdYear = intRocYear + ROC_YEAR_OFFSET
    
    ' 西元閏年規則：被 4 整除且不能被 100 整除，或能被 400 整除
    If (intAdYear Mod 4 = 0 And intAdYear Mod 100 <> 0) Or (intAdYear Mod 400 = 0) Then
        IsRocLeapYear = True
    Else
        IsRocLeapYear = False
    End If
End Function

''' <summary>
''' 將民國日期字串（支援 7碼 YYYMMDD, 6碼 YYMMDD, YYY-MM-DD, YYY/MM/DD）轉換為 VBA Date 物件
''' </summary>
Public Function CDateFromROC(ByVal strRocDate As String) As Date
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    Dim i As Integer
    Dim ch As String
    
    ' 1. 清理字串，只留下數字與分隔符（/ 或 -）
    strRocDate = Trim(strRocDate)
    strClean = ""
    For i = 1 To Len(strRocDate)
        ch = Mid(strRocDate, i, 1)
        If IsNumeric(ch) Or ch = "/" Or ch = "-" Then
            strClean = strClean & ch
        End If
    Next i
    
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim parts() As String
    
    ' 2. 判斷是否帶有分隔符
    If InStr(strClean, "/") > 0 Then
        parts = Split(strClean, "/")
    ElseIf InStr(strClean, "-") > 0 Then
        parts = Split(strClean, "-")
    End If
    
    If IsArrayInitialized(parts) Then
        If UBound(parts) = 2 Then
            intYear = CInt(parts(0))
            intMonth = CInt(parts(1))
            intDay = CInt(parts(2))
        Else
            Err.Raise vbObjectError + 1001, "CDateFromROC", "日期格式不符拆分標準"
        End If
    Else
        ' 無分隔符，處理 7 碼 (YYYMMDD) 或 6 碼 (YYMMDD)
        If Len(strClean) = 7 Then
            intYear = CInt(Left(strClean, 3))
            intMonth = CInt(Mid(strClean, 4, 2))
            intDay = CInt(Right(strClean, 2))
        ElseIf Len(strClean) = 6 Then
            intYear = CInt(Left(strClean, 2))
            intMonth = CInt(Mid(strClean, 3, 2))
            intDay = CInt(Right(strClean, 2))
        Else
            Err.Raise vbObjectError + 1002, "CDateFromROC", "純數字日期長度非 6 或 7 碼"
        End If
    End If
    
    ' 3. 月份與日期合理性檢查
    If intMonth < 1 Or intMonth > 12 Then
        Err.Raise vbObjectError + 1003, "CDateFromROC", "月份超出範圍 (1-12)"
    End If
    
    Dim intMaxDays As Integer
    intMaxDays = GetDaysInMonth(intYear, intMonth)
    If intDay < 1 Or intDay > intMaxDays Then
        Err.Raise vbObjectError + 1004, "CDateFromROC", "日期超出該月天數限制 (" & intMaxDays & "天)"
    End If
    
    ' 4. 轉為西元日期
    CDateFromROC = DateSerial(intYear + ROC_YEAR_OFFSET, intMonth, intDay)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "modUtility.CDateFromROC", "民國日期「" & strRocDate & "」轉換失敗: " & Err.Description
End Function

''' <summary>
''' 解析民國日期時間字串（例如 "115-02-10 08:00"）為 VBA Date 物件
''' </summary>
Public Function ParseRocDateTime(ByVal strRocDateTime As String) As Date
    On Error GoTo ErrorHandler
    
    strRocDateTime = Trim(strRocDateTime)
    Dim spacePos As Integer
    spacePos = InStr(strRocDateTime, " ")
    
    If spacePos = 0 Then
        ' 無時間部分，只解析日期
        ParseRocDateTime = CDateFromROC(strRocDateTime)
        Exit Function
    End If
    
    Dim strDatePart As String
    Dim strTimePart As String
    strDatePart = Left(strRocDateTime, spacePos - 1)
    strTimePart = Trim(Mid(strRocDateTime, spacePos + 1))
    
    ' 解析日期部分
    Dim dtDate As Date
    dtDate = CDateFromROC(strDatePart)
    
    ' 解析時間部分 (HH:MM)
    Dim timeParts() As String
    timeParts = Split(strTimePart, ":")
    
    Dim intHour As Integer, intMin As Integer
    intHour = CInt(timeParts(0))
    intMin = CInt(timeParts(1))
    
    If intHour < 0 Or intHour > 23 Or intMin < 0 Or intMin > 59 Then
        Err.Raise vbObjectError + 1005, "ParseRocDateTime", "時間格式超出範圍 (00:00-23:59)"
    End If
    
    ' 合成完整 Date 物件
    ParseRocDateTime = DateAdd("h", intHour, dtDate)
    ParseRocDateTime = DateAdd("n", intMin, ParseRocDateTime)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "modUtility.ParseRocDateTime", "民國日期時間「" & strRocDateTime & "」轉換失敗: " & Err.Description
End Function

''' <summary>
''' 將西元 Date 轉為民國日期字串 (預設格式 YYY/MM/DD)
''' </summary>
Public Function FormatROC(ByVal dtDate As Date, Optional ByVal strFormat As String = "YYY/MM/DD") As String
    Dim intAdYear As Integer
    Dim intRocYear As Integer
    Dim strMonth As String
    Dim strDay As String
    
    intAdYear = Year(dtDate)
    intRocYear = intAdYear - ROC_YEAR_OFFSET
    strMonth = Format(Month(dtDate), "00")
    strDay = Format(Day(dtDate), "00")
    
    Dim strRocYear As String
    strRocYear = Format(intRocYear, "000")
    
    Dim result As String
    result = strFormat
    result = Replace(result, "YYY", strRocYear)
    result = Replace(result, "MM", strMonth)
    result = Replace(result, "DD", strDay)
    
    FormatROC = result
End Function

' --- 內部輔助函式 ---

Private Function IsArrayInitialized(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim lBoundCheck As Long
    lBoundCheck = LBound(arr)
    If Err.Number = 0 Then
        IsArrayInitialized = True
    Else
        IsArrayInitialized = False
    End If
End Function

Private Function GetDaysInMonth(ByVal intRocYear As Integer, ByVal intMonth As Integer) As Integer
    Select Case intMonth
        Case 1, 3, 5, 7, 8, 10, 12
            GetDaysInMonth = 31
        Case 4, 6, 9, 11
            GetDaysInMonth = 30
        Case 2
            If IsRocLeapYear(intRocYear) Then
                GetDaysInMonth = 29
            Else
                GetDaysInMonth = 28
            End If
    End Select
End Function
