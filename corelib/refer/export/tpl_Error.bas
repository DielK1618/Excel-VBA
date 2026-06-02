Attribute VB_Name = "tpl_Error"
Option Explicit
' 로그 설정 (True: 로그 기록, False: 로그 기록 안함)
Public Const ENABLE_ERROR_LOG As Boolean = False  ' 테스트 시 False
'Public Const ENABLE_ERROR_LOG As Boolean = True   ' 운영 시 True

' 에러 핸들러
Public Sub HandleError(Optional procName As String = "", _
                      Optional additionalInfo As String = "", _
                      Optional showMessage As Boolean = True)
    
    Dim errMsg As String
    Dim logMsg As String
    
    ' 화면 표시용 메시지
    errMsg = "오류가 발생했습니다." & vbCrLf & vbCrLf
    If procName <> "" Then errMsg = errMsg & "프로시저: " & procName & vbCrLf
    errMsg = errMsg & "오류 번호: " & Err.Number & vbCrLf
    errMsg = errMsg & "오류 내용: " & Err.Description
    If additionalInfo <> "" Then errMsg = errMsg & vbCrLf & "추가 정보: " & additionalInfo
    
    If showMessage Then
        MsgBox errMsg, vbCritical, "오류"
    End If
    
    ' 전역 설정에 따라 로그 기록
    If ENABLE_ERROR_LOG Then
        logMsg = procName & " | " & _
                Err.Number & " | " & _
                Err.Description & " | " & _
                additionalInfo
        WriteLog logMsg, "ERROR"
    End If
    
    cl.sht_Lock
    cl.Calculate_On
    cl.Event_On
    cl.DPUpdate_On
    
End Sub

' 로그 기록
Public Sub WriteLog(logMessage As String, Optional logType As String = "INFO")
    ' 전역 설정 확인
    If Not ENABLE_ERROR_LOG Then Exit Sub
    
    Dim logFile As String
    Dim fileNum As Integer
    Dim logFolder As String
    
    On Error Resume Next
    
    logFolder = ThisWorkbook.Path & "\Logs\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder
    
    logFile = logFolder & "Log_" & Format(Date, "yyyymm") & ".txt"
    
    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | [" & logType & "] | " & logMessage
    Close #fileNum
    
End Sub
