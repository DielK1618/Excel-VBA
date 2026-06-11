Attribute VB_Name = "am_Core"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Core                                                │
' │  역할 : 전역 상수, 프로퍼티, 초기화/정리, 로드 확인,   │
' │         Application 상태 관리, 워크북 보호              │
' └─────────────────────────────────────────────────────────┘

' ── 1. 전역 상수 선언 ────────────────────────────────────────
Public Const AM_VERSION   As String = "1.0.0"
Public Const AM_NAME      As String = "corelib"
Public Const AM_FILE      As String = "corelib.xlam"
Public Const CM_TO_POINTS As Double = 28.3465   ' cm → Point 단위 변환

' ── 2. 모듈 변수 선언 ────────────────────────────────────────
Private m_blnReady As Boolean

' ── 3. 모듈 프로퍼티 ─────────────────────────────────────────

' 목적   : xlam 폴더 경로 반환
' 반환   : String - 예) "D:\Tools"
' 예시   : am_Core.XlamPath → "D:\Tools"
Public Function XlamPath() As String
    XlamPath = ThisWorkbook.Path
End Function

' 목적   : xlam 전체 파일 경로 반환
' 반환   : String - 예) "D:\Tools\corelib.xlam"
' 예시   : am_Core.XlamFullName → "D:\Tools\corelib.xlam"
Public Function XlamFullName() As String
    XlamFullName = ThisWorkbook.FullName
End Function

' 목적   : xlam 버전 반환
' 반환   : String - 예) "1.0.0"
' 예시   : am_Core.Version → "1.0.0"
Public Function Version() As String
    Version = AM_VERSION
End Function

' 목적   : xlam 초기화 완료 상태 반환
' 반환   : Boolean - True: 정상 초기화 / False: 미초기화
' 예시   : If am_Core.IsReady Then ...
Public Function IsReady() As Boolean
    If Not m_blnReady Then Call Initialize()
    IsReady = m_blnReady
End Function

' ── 4. 초기화 / 정리 ─────────────────────────────────────────

' 목적   : xlam 열림 시 전역 리소스 초기화
' 호출   : ThisWorkbook.Workbook_Open 에서 자동 호출
Public Sub Initialize()

    On Error GoTo ErrHandler

    m_blnReady = False
    m_blnReady = True
    Exit Sub

ErrHandler:
    m_blnReady = False
    MsgBox "[" & AM_NAME & "] 초기화 실패" & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, AM_NAME

End Sub

' 목적   : xlam 닫힘 시 리소스 정리
' 호출   : ThisWorkbook.Workbook_BeforeClose 에서 자동 호출
Public Sub CleanUp()

    On Error Resume Next
    m_blnReady = False

End Sub

' ── 5. xlam 로드 확인 ────────────────────────────────────────

' 목적   : corelib.xlam 이 현재 열려있는지 확인
' 반환   : Boolean - True: 로드됨 / False: 미로드
' 참고   : Workbooks 컬렉션에 xlam 이 포함되지 않는 특성상
'          오류 방식으로 체크
' 예시   : If am_Core.IsXlamLoaded Then ...
Public Function IsXlamLoaded() As Boolean

    Dim wbXlam  As Workbook
    Dim strTest As String

    On Error Resume Next
    Set wbXlam = Workbooks(AM_FILE)
    strTest = wbXlam.Name
    On Error GoTo 0

    IsXlamLoaded = (strTest <> "")

End Function

' ══════════════════════════════════════════════════════════
'  Application 상태 관리
' ══════════════════════════════════════════════════════════

' ── 6. 화면 업데이트 / 경고창 ────────────────────────────────

' 목적   : 화면 갱신·경고창 비활성화 (작업 시작 전 호출)
' 예시   : am_Core.DPUpdate_Off
Public Sub DPUpdate_Off()
    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
End Sub

' 목적   : 화면 갱신·경고창 활성화 복원 (작업 완료 후 호출)
' 예시   : am_Core.DPUpdate_On
Public Sub DPUpdate_On()
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
End Sub

' ── 7. 이벤트 ────────────────────────────────────────────────

' 목적   : 이벤트 비활성화
' 예시   : am_Core.Event_Off
Public Sub Event_Off()
    Application.EnableEvents = False
End Sub

' 목적   : 이벤트 활성화 복원
' 예시   : am_Core.Event_On
Public Sub Event_On()
    Application.EnableEvents = True
End Sub

' ── 8. 계산 모드 ─────────────────────────────────────────────

' 목적   : 자동 계산 중단 (수동 모드 전환)
' 예시   : am_Core.Calculate_Off
Public Sub Calculate_Off()
    Application.Calculation = xlCalculationManual
End Sub

' 목적   : 자동 계산 복원
' 예시   : am_Core.Calculate_On
Public Sub Calculate_On()
    Application.Calculation = xlCalculationAutomatic
End Sub

' ══════════════════════════════════════════════════════════
'  워크북 보호
' ══════════════════════════════════════════════════════════

' ── 9. 워크북 구조 보호 ──────────────────────────────────────

' 목적   : 워크북 구조(시트 추가·삭제·이동) 보호
' 인수   : wb    - 보호할 워크북 (기본: ActiveWorkbook)
'          strPW - 보호 비밀번호 (기본: 없음)
' 예시   : am_Core.WB_Lock ThisWorkbook, "1234"
Public Sub WB_Lock(Optional ByVal wb  As Workbook, _
                   Optional ByVal strPW As String = "")

    If wb Is Nothing Then Set wb = ActiveWorkbook
    wb.Protect Password:=strPW, Structure:=True

End Sub

' 목적   : 워크북 구조 보호 해제
' 인수   : wb    - 해제할 워크북 (기본: ActiveWorkbook)
'          strPW - 보호 비밀번호 (기본: 없음)
' 예시   : am_Core.WB_UnLock ThisWorkbook, "1234"
Public Sub WB_UnLock(Optional ByVal wb  As Workbook, _
                     Optional ByVal strPW As String = "")

    If wb Is Nothing Then Set wb = ActiveWorkbook
    wb.Unprotect Password:=strPW

End Sub
