Attribute VB_Name = "ref_Automation"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Automation                                          │
' │  역할 : am_Automation Application.Run 래퍼               │
' └─────────────────────────────────────────────────────────┘

' 참고 : 아래 프로시저는 래핑 불가 → 직접 Application.Run 사용
'        - ExecuteKeySequence          : ParamArray 파라미터
'        - ProcessRangeWithKeySequence : ParamArray 파라미터
'        - GetMousePosition            : ByRef Long 파라미터

Private Const REF As String = "corelib.xlam!am_Automation."

' ── Enum 재선언 (CWB 직접 참조용) ────────────────────────────
' xlam 에 정의된 Public Enum 은 CWB 에서 이름으로 접근 불가
' → ref_Automation 에 동일 값으로 재선언하여 CWB 코드에서 상수처럼 사용

Public Enum VirtualKeys
    VK_TAB     = &H9
    VK_RETURN  = &HD
    VK_CONTROL = &H11
    VK_MENU    = &H12
    VK_ESCAPE  = &H1B
    VK_SPACE   = &H20
    VK_LEFT    = &H25
    VK_UP      = &H26
    VK_RIGHT   = &H27
    VK_DOWN    = &H28
End Enum

Public Enum KeyActions
    ACTION_COPY        = 1
    ACTION_PASTE       = 2
    ACTION_TAB         = 3
    ACTION_ENTER       = 4
    ACTION_ALT_TAB     = 5
    ACTION_ESCAPE      = 6
    ACTION_ARROW_DOWN  = 7
    ACTION_ARROW_UP    = 8
    ACTION_ARROW_LEFT  = 9
    ACTION_ARROW_RIGHT = 10
End Enum

' ── 키보드 ───────────────────────────────────────────────────

' 목적   : 단일 키보드 액션 실행 (KeyActions Enum 사용)
' 인수   : action    - 실행할 액션 (KeyActions Enum: ACTION_COPY 등)
'          waitAfter - 액션 후 대기 시간(ms, 기본: 100)
' 예시   : ExecuteKeyAction ACTION_COPY
'          ExecuteKeyAction ACTION_PASTE, 200
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Application.Run REF & "ExecuteKeyAction", CLng(action), waitAfter
End Sub

' ── 마우스 ───────────────────────────────────────────────────

' 목적   : 화면 절대 좌표 위치 클릭
' 인수   : lngX      - 클릭 X 좌표 (픽셀)
'          lngY      - 클릭 Y 좌표 (픽셀)
'          blnLeft   - True: 좌클릭 / False: 우클릭 (기본: True)
'          strDelay  - 클릭 전 대기 시간 문자열 (예: "00:00:01")
'          strWait   - 클릭 후 대기 시간 문자열
' 예시   : ClickAtPosition 500, 300
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    Application.Run REF & "ClickAtPosition", lngX, lngY, blnLeft, strDelay, strWait
End Sub

' 목적   : 특정 시각까지 대기 (Application.Wait 래퍼)
' 인수   : strTime - 대기 종료 시각 문자열 (예: "00:00:02" → 2초 후)
' 예시   : WaitTime "00:00:03"
Public Sub WaitTime(ByVal strTime As String)
    Application.Run REF & "WaitTime", strTime
End Sub

' ── 창 제어 ───────────────────────────────────────────────────

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 SendMessage(BM_CLICK)로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
' 반환   : Boolean - 버튼을 찾아 클릭 메시지를 보냈는지 여부
' 예시   : ClickButtonByText("WMC SCDK", "확인")
Public Function ClickButtonByText(ByVal strWindowTitle As String, ByVal strButtonText As String) As Boolean
    ClickButtonByText = Application.Run(REF & "ClickButtonByText", strWindowTitle, strButtonText)
End Function

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 중심 좌표를 마우스로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
'          strClickDelay  - 좌표 클릭 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 버튼을 찾아 클릭했는지 여부
Public Function ClickButtonByRect(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                  Optional ByVal strClickDelay As String = "00:00:02") As Boolean
    ClickButtonByRect = Application.Run(REF & "ClickButtonByRect", strWindowTitle, strButtonText, strClickDelay)
End Function

' 목적   : ClickButtonByRect 재시도 래퍼 (반복문 등 타이밍이 불안정한 구간용)
' 인수   : strWindowTitle - 대상 창 제목
'          strButtonText  - 찾을 버튼 텍스트
'          lngMaxRetry    - 최대 재시도 횟수 (기본 3)
'          strRetryWait   - 실패 시 재시도 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 클릭 성공 여부
Public Function ClickButtonByRect_Retry(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                        Optional ByVal lngMaxRetry As Long = 3, _
                                        Optional ByVal strRetryWait As String = "00:00:02") As Boolean
    ClickButtonByRect_Retry = Application.Run(REF & "ClickButtonByRect_Retry", strWindowTitle, strButtonText, lngMaxRetry, strRetryWait)
End Function

' 목적   : 다른 프로세스 소유의 창을 강제로 포그라운드로 활성화
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
' 반환   : Boolean - 창을 찾아 활성화를 시도했는지 여부
Public Function ForceActivateWindow(ByVal strWindowTitle As String) As Boolean
    ForceActivateWindow = Application.Run(REF & "ForceActivateWindow", strWindowTitle)
End Function

' 목적   : UI Automation(Late Binding)으로 창 안의 버튼을 이름으로 찾아 Invoke
' 인수   : strWindowName - 대상 창 이름 (UIA Name 속성 기준)
'          strButtonName - 찾을 버튼 이름 (UIA Name 속성 기준)
' 반환   : Boolean - Invoke 성공 여부
Public Function ClickButtonByUIA(ByVal strWindowName As String, ByVal strButtonName As String) As Boolean
    ClickButtonByUIA = Application.Run(REF & "ClickButtonByUIA", strWindowName, strButtonName)
End Function
