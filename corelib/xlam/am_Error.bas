Attribute VB_Name = "am_Error"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Error                                               │
' │  역할 : 공통 에러 핸들링, 로그 파일 기록               │
' └─────────────────────────────────────────────────────────┘

' ── 1. 모듈 변수 선언 ─────────────────────────────────────────
Private m_blnLogEnabled As Boolean   ' 기본: False (운영 시 SetLogEnabled True 호출)

' ══════════════════════════════════════════════════════════
'  로그 활성화 제어
' ══════════════════════════════════════════════════════════

' 목적   : 로그 기록 활성화 여부 설정
' 인수   : blnEnabled - True: 로그 기록 / False: 무동작 (기본)
' 예시   : am_Error.SetLogEnabled True
Public Sub SetLogEnabled(ByVal blnEnabled As Boolean)
    m_blnLogEnabled = blnEnabled
End Sub

' 목적   : 로그 기록 활성화 여부 반환
' 반환   : Boolean - True: 활성 / False: 비활성
' 예시   : If am_Error.GetLogEnabled Then ...
Public Function GetLogEnabled() As Boolean
    GetLogEnabled = m_blnLogEnabled
End Function

' ══════════════════════════════════════════════════════════
'  에러 핸들링
' ══════════════════════════════════════════════════════════

' 목적   : 공통 에러 핸들러 — 각 모듈 ErrHandler 레이블에서 호출
' 인수   : strProcName    - 오류 발생 프로시저명 (기본: "")
'          strAdditional  - 추가 정보 메시지    (기본: "")
'          blnShowMessage - 오류 메시지 표시 여부 (기본: True)
' 예시   : am_Error.HandleError "am_File.MkFolder", "대상 경로 없음"
Public Sub HandleError(Optional ByVal strProcName    As String  = "", _
                       Optional ByVal strAdditional  As String  = "", _
                       Optional ByVal blnShowMessage As Boolean = True)

    Dim strErrMsg As String
    Dim strLogMsg As String

    strErrMsg = "오류가 발생했습니다." & vbCrLf & vbCrLf
    If strProcName <> "" Then strErrMsg = strErrMsg & "프로시저: " & strProcName & vbCrLf
    strErrMsg = strErrMsg & "오류 번호: " & Err.Number & vbCrLf
    strErrMsg = strErrMsg & "오류 내용: " & Err.Description
    If strAdditional <> "" Then strErrMsg = strErrMsg & vbCrLf & "추가 정보: " & strAdditional

    If blnShowMessage Then
        MsgBox strErrMsg, vbCritical, am_Core.AM_NAME
    End If

    If m_blnLogEnabled Then
        strLogMsg = strProcName & " | " & _
                    Err.Number & " | " & _
                    Err.Description & " | " & _
                    strAdditional
        WriteLog strLogMsg, "ERROR"
    End If

End Sub

' ── 2. 로그 파일 기록 ─────────────────────────────────────────

' 목적   : xlam 폴더 하위 Logs 디렉터리에 월별 로그 파일 기록
' 인수   : strMessage - 기록할 메시지
'          strType    - 로그 유형 (기본: "INFO")
' 예시   : am_Error.WriteLog "초기화 완료", "INFO"
Public Sub WriteLog(ByVal strMessage As String, _
                    Optional ByVal strType As String = "INFO")

    If Not m_blnLogEnabled Then Exit Sub

    Dim strLogFolder As String
    Dim strLogFile   As String
    Dim intFileNum   As Integer

    On Error Resume Next

    strLogFolder = am_Core.XlamPath & "\Logs\"
    If Dir(strLogFolder, vbDirectory) = "" Then MkDir strLogFolder

    strLogFile = strLogFolder & "Log_" & Format(Date, "yyyymm") & ".txt"

    intFileNum = FreeFile
    Open strLogFile For Append As #intFileNum
    Print #intFileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | [" & strType & "] | " & strMessage
    Close #intFileNum

End Sub
