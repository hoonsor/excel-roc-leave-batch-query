Attribute VB_Name = "modUtility"
Option Explicit

' ¢Dx?W????¢D?¢X????P????¢X??t??¡Ó`??
Public Const ROC_YEAR_OFFSET As Integer = 1911

''' <summary>
''' ¡±P?_¢D?¢X??~?O¡±_?¢X?|?~
''' </summary>
Public Function IsRocLeapYear(ByVal intRocYear As Integer) As Boolean
    Dim intAdYear As Integer
    intAdYear = intRocYear + ROC_YEAR_OFFSET
    
    ' ?????|?~?W?h?G?Q 4 ??¢X¢G¢DB?¢G¡Â??Q 100 ??¢X¢G?A??¡Â??Q 400 ??¢X¢G
    If (intAdYear Mod 4 = 0 And intAdYear Mod 100 <> 0) Or (intAdYear Mod 400 = 0) Then
        IsRocLeapYear = True
    Else
        IsRocLeapYear = False
    End If
End Function

''' <summary>
''' ¡ÓN¢D?¢X??????r???]???? 7?X YYYMMDD, 6?X YYMMDD, YYY-MM-DD, YYY/MM/DD?^?????¢X VBA Date ??¢D?
''' </summary>
Public Function CDateFromROC(ByVal strRocDate As String) As Date
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    Dim i As Integer
    Dim ch As String
    
    ' 1. ?M?z?r???A¢Du¡Âd?U???r?P???j???]/ ?? -?^
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
    
    ' 2. ¡±P?_?O¡±_¡Óa?????j??
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
            Err.Raise vbObjectError + 1001, "CDateFromROC", "?????????¢G????????¡P?"
        End If
    Else
        ' ?L???j???A?B?z 7 ?X (YYYMMDD) ?? 6 ?X (YYMMDD)
        If Len(strClean) = 7 Then
            intYear = CInt(Left(strClean, 3))
            intMonth = CInt(Mid(strClean, 4, 2))
            intDay = CInt(Right(strClean, 2))
        ElseIf Len(strClean) = 6 Then
            intYear = CInt(Left(strClean, 2))
            intMonth = CInt(Mid(strClean, 3, 2))
            intDay = CInt(Right(strClean, 2))
        Else
            Err.Raise vbObjectError + 1002, "CDateFromROC", "¡Â????r???????¡Ñ?D 6 ?? 7 ?X"
        End If
    End If
    
    ' 3. ??¢D¡Ò?P?????X?z?????d
    If intMonth < 1 Or intMonth > 12 Then
        Err.Raise vbObjectError + 1003, "CDateFromROC", "??¢D¡Ò?W¢DX?d?? (1-12)"
    End If
    
    Dim intMaxDays As Integer
    intMaxDays = GetDaysInMonth(intYear, intMonth)
    If intDay < 1 Or intDay > intMaxDays Then
        Err.Raise vbObjectError + 1004, "CDateFromROC", "?????W¢DX???????????? (" & intMaxDays & "??)"
    End If
    
    ' 4. ???¢X????????
    CDateFromROC = DateSerial(intYear + ROC_YEAR_OFFSET, intMonth, intDay)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "modUtility.CDateFromROC", "¢D?¢X??????u" & strRocDate & "?v????¢D¢F¡Ó?: " & Err.Description
End Function

''' <summary>
''' ???R¢D?¢X??????????r???]???p "115-02-10 08:00"?^?¢X VBA Date ??¢D?
''' </summary>
Public Function ParseRocDateTime(ByVal strRocDateTime As String) As Date
    On Error GoTo ErrorHandler
    
    strRocDateTime = Trim(strRocDateTime)
    Dim spacePos As Integer
    spacePos = InStr(strRocDateTime, " ")
    
    If spacePos = 0 Then
        ' ?L?????????A¢Du???R????
        ParseRocDateTime = CDateFromROC(strRocDateTime)
        Exit Function
    End If
    
    Dim strDatePart As String
    Dim strTimePart As String
    strDatePart = Left(strRocDateTime, spacePos - 1)
    strTimePart = Trim(Mid(strRocDateTime, spacePos + 1))
    
    ' ???R????????
    Dim dtDate As Date
    dtDate = CDateFromROC(strDatePart)
    
    ' ???R???????? (HH:MM)
    Dim timeParts() As String
    timeParts = Split(strTimePart, ":")
    
    Dim intHour As Integer, intMin As Integer
    intHour = CInt(timeParts(0))
    intMin = CInt(timeParts(1))
    
    If intHour < 0 Or intHour > 23 Or intMin < 0 Or intMin > 59 Then
        Err.Raise vbObjectError + 1005, "ParseRocDateTime", "?????????W¢DX?d?? (00:00-23:59)"
    End If
    
    ' ?X??¡±??? Date ??¢D?
    ParseRocDateTime = DateAdd("h", intHour, dtDate)
    ParseRocDateTime = DateAdd("n", intMin, ParseRocDateTime)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "modUtility.ParseRocDateTime", "¢D?¢X??????????u" & strRocDateTime & "?v????¢D¢F¡Ó?: " & Err.Description
End Function

''' <summary>
''' ¡ÓN???? Date ???¢X¢D?¢X??????r?? (?w?]???? YYY/MM/DD)
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

' --- ??????¡±U???? ---

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
