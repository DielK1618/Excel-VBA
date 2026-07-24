Attribute VB_Name = "am_Automation"
Option Explicit

' +---------------------------------------------------------+
' |  am_Automation                                          |
' |  역할 : 키보드/마우스 입력 시뮬레이션, 창 탐색/활성화/버튼 제어 |
' +---------------------------------------------------------+

' -- Windows API --
Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" _
        (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function EnumChildWindows Lib "user32" _
        (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowTextA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function SendMessageA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" _
        (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" _
        (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    lngLeft   As Long
    lngTop    As Long
    lngRight  As Long
    lngBottom As Long
End Type

Private Const KEYEVENTF_KEYDOWN     As Long = 0
Private Const KEYEVENTF_KEYUP       As Long = 2
Private Const MOUSEEVENTF_LEFTDOWN  As Long = &H2
Private Const MOUSEEVENTF_LEFTUP    As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

Private Const BM_CLICK   As Long = &HF5
Private Const SW_RESTORE As Long = 9

' UI Automation 속성/패턴 ID — Late Binding(CreateObject) 방식이라 참조 추가 없이 그대로 사용
Private Const UIA_NamePropertyId        As Long = 30005
Private Const UIA_InvokePatternId       As Long = 10000
Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_ButtonControlTypeId   As Long = 50000
Private Const TreeScope_Descendants     As Long = 4

' 버튼 탐색 결과 임시 저장 (prv_EnumChildProc 콜백 전용)
Private m_hWndFound     As LongPtr
Private m_strTargetText As String

' 가상 키코드
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

' 키 동작 정의
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

' ==========================================================
'  키보드
' ==========================================================

' 목적   : KeyActions Enum에 정의된 키 동작 실행
' 인수   : action    - 실행할 키 동작 (KeyActions Enum 또는 Long)
'          waitAfter - 실행 후 대기 ms (기본 100)
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Select Case action
        Case ACTION_COPY        : prv_SendKeyCombo VK_CONTROL, 67   ' Ctrl+C
        Case ACTION_PASTE       : prv_SendKeyCombo VK_CONTROL, 86   ' Ctrl+V
        Case ACTION_TAB         : prv_SendKey VK_TAB
        Case ACTION_ENTER       : prv_SendKey VK_RETURN
        Case ACTION_ALT_TAB     : prv_SendKeyCombo VK_MENU, VK_TAB
        Case ACTION_ESCAPE      : prv_SendKey VK_ESCAPE
        Case ACTION_ARROW_DOWN  : prv_SendKey VK_DOWN
        Case ACTION_ARROW_UP    : prv_SendKey VK_UP
        Case ACTION_ARROW_LEFT  : prv_SendKey VK_LEFT
        Case ACTION_ARROW_RIGHT : prv_SendKey VK_RIGHT
    End Select

    If waitAfter > 0 Then prv_WaitMs waitAfter
End Sub

' 목적   : 여러 키 동작을 순서대로 실행
' 인수   : actions - KeyActions 또는 Array(KeyActions, waitMs) 의 ParamArray
Public Sub ExecuteKeySequence(ParamArray actions() As Variant)
    Dim i As Long

    For i = LBound(actions) To UBound(actions)
        If IsArray(actions(i)) Then
            ExecuteKeyAction actions(i)(0), actions(i)(1)
        Else
            ExecuteKeyAction actions(i)
        End If
    Next i
End Sub

' 목적   : 범위의 각 셀에 키 시퀀스 적용
' 인수   : rng     - 대상 범위
'          actions - ExecuteKeySequence에 전달할 동작 ParamArray
Public Sub ProcessRangeWithKeySequence(ByVal rng As Range, _
                                       ParamArray actions() As Variant)
    Dim cel As Range

    For Each cel In rng
        If Not IsEmpty(cel) Then
            cel.Select
            ExecuteKeySequence actions
        End If
    Next cel
End Sub

Private Sub prv_SendKey(ByVal bKeyCode As Byte, _
                        Optional ByVal blnRelease As Boolean = True)
    keybd_event bKeyCode, 0, KEYEVENTF_KEYDOWN, 0
    If blnRelease Then keybd_event bKeyCode, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub prv_SendKeyCombo(ParamArray keyCodes() As Variant)
    Dim i As Long

    For i = LBound(keyCodes) To UBound(keyCodes)
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYDOWN, 0
    Next i

    prv_WaitMs 50

    For i = UBound(keyCodes) To LBound(keyCodes) Step -1
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYUP, 0
    Next i
End Sub

' 목적   : ms 단위 대기 (키보드/마우스 내부 공용)
Private Sub prv_WaitMs(ByVal lngMs As Long)
    Application.Wait Now + TimeSerial(0, 0, lngMs / 1000)
End Sub

' ==========================================================
'  마우스
' ==========================================================

' 목적   : 현재 마우스 커서 위치 반환
' 인수   : lngX - (반환) X 좌표
'          lngY - (반환) Y 좌표
Public Sub GetMousePosition(ByRef lngX As Long, ByRef lngY As Long)
    Dim pt As POINTAPI
    GetCursorPos pt
    lngX = pt.x
    lngY = pt.y
End Sub

' 목적   : 마우스를 지정 좌표로 이동 후 클릭
' 인수   : lngX     - X 좌표
'          lngY     - Y 좌표
'          blnLeft  - True=왼쪽 클릭, False=오른쪽 클릭 (기본 True)
'          strDelay - 이동 후 클릭 전 대기 시간 "HH:MM:SS" (생략 시 생략)
'          strWait  - 클릭 후 대기 시간 "HH:MM:SS" (생략 시 생략)
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    SetCursorPos lngX, lngY

    If strDelay <> "" Then WaitTime strDelay

    If blnLeft Then
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Else
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    End If

    If strWait <> "" Then WaitTime strWait
End Sub

' 목적   : 입력된 시간만큼 대기 (Application.Wait 기반)
' 인수   : strTime - 대기 시간 "HH:MM:SS" 형식 (예: "00:00:03" = 3초)
Public Sub WaitTime(ByVal strTime As String)
    Application.Wait Now + TimeValue(strTime)
End Sub

' ==========================================================
'  창 제어
' ==========================================================

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 SendMessage(BM_CLICK)로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
' 반환   : Boolean - 버튼을 찾아 클릭 메시지를 보냈는지 여부
' 예시   : ClickButtonByText("WMC SCDK", "확인")
Public Function ClickButtonByText(ByVal strWindowTitle As String, ByVal strButtonText As String) As Boolean
    Dim hWndParent As LongPtr

    hWndParent = FindWindow(vbNullString, strWindowTitle)
    If hWndParent = 0 Then
        ClickButtonByText = False
        Exit Function
    End If

    m_hWndFound = 0
    m_strTargetText = strButtonText
    EnumChildWindows hWndParent, AddressOf prv_EnumChildProc, 0

    If m_hWndFound = 0 Then
        ClickButtonByText = False
        Exit Function
    End If

    SendMessageA m_hWndFound, BM_CLICK, 0, 0
    ClickButtonByText = True
End Function

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 중심 좌표를 마우스로 클릭
'          (SendMessage(BM_CLICK)가 통하지 않는 커스텀/논표준 버튼용)
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
'          strClickDelay  - 좌표 클릭 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 버튼을 찾아 클릭했는지 여부
Public Function ClickButtonByRect(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                  Optional ByVal strClickDelay As String = "00:00:02") As Boolean
    Dim hWndParent As LongPtr
    Dim udtRect    As RECT
    Dim lngCenterX As Long, lngCenterY As Long

    hWndParent = FindWindow(vbNullString, strWindowTitle)
    If hWndParent = 0 Then
        ClickButtonByRect = False
        Exit Function
    End If

    m_hWndFound = 0
    m_strTargetText = strButtonText
    EnumChildWindows hWndParent, AddressOf prv_EnumChildProc, 0

    If m_hWndFound = 0 Then
        ClickButtonByRect = False
        Exit Function
    End If

    GetWindowRect m_hWndFound, udtRect
    lngCenterX = (udtRect.lngLeft + udtRect.lngRight) \ 2
    lngCenterY = (udtRect.lngTop + udtRect.lngBottom) \ 2

    ClickAtPosition lngCenterX, lngCenterY, strDelay:=strClickDelay
    ClickButtonByRect = True
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
    Dim lngRetryCount As Long

    For lngRetryCount = 1 To lngMaxRetry
        If ClickButtonByRect(strWindowTitle, strButtonText) Then
            ClickButtonByRect_Retry = True
            Exit Function
        End If
        WaitTime strRetryWait
    Next lngRetryCount

    ClickButtonByRect_Retry = False
End Function

' 목적   : 다른 프로세스 소유의 창을 강제로 포그라운드로 활성화 (AttachThreadInput 사용)
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
' 반환   : Boolean - 창을 찾아 활성화를 시도했는지 여부
Public Function ForceActivateWindow(ByVal strWindowTitle As String) As Boolean
    Dim hWndTarget          As LongPtr, hWndForeground As LongPtr
    Dim lngForegroundThread As Long, lngCurrentThread As Long
    Dim lngDummy             As Long

    hWndTarget = FindWindow(vbNullString, strWindowTitle)
    If hWndTarget = 0 Then
        ForceActivateWindow = False
        Exit Function
    End If

    hWndForeground = GetForegroundWindow()
    lngForegroundThread = GetWindowThreadProcessId(hWndForeground, lngDummy)
    lngCurrentThread = GetCurrentThreadId()

    AttachThreadInput lngCurrentThread, lngForegroundThread, True
    ShowWindow hWndTarget, SW_RESTORE
    BringWindowToTop hWndTarget
    SetForegroundWindow hWndTarget
    AttachThreadInput lngCurrentThread, lngForegroundThread, False

    ForceActivateWindow = True
End Function

' 목적   : UI Automation(Late Binding)으로 창 안의 버튼을 이름으로 찾아 Invoke
' 인수   : strWindowName - 대상 창 이름 (UIA Name 속성 기준)
'          strButtonName - 찾을 버튼 이름 (UIA Name 속성 기준)
' 반환   : Boolean - Invoke 성공 여부
' 참고   : CreateObject("UIAutomationClient.CUIAutomation") 사용 — 프로젝트 참조 추가 불필요 (Late Binding, am_ 설계 원칙 준수)
Public Function ClickButtonByUIA(ByVal strWindowName As String, ByVal strButtonName As String) As Boolean
    On Error GoTo ErrHandler

    Dim oAutomation    As Object
    Dim oRoot          As Object
    Dim oCondition      As Object
    Dim oWindowElement  As Object
    Dim oButtonElement  As Object
    Dim oInvokePattern  As Object

    Set oAutomation = CreateObject("UIAutomationClient.CUIAutomation")
    Set oRoot = oAutomation.GetRootElement

    Set oCondition = oAutomation.CreatePropertyCondition(UIA_NamePropertyId, strWindowName)
    Set oWindowElement = oRoot.FindFirst(TreeScope_Descendants, oCondition)

    If oWindowElement Is Nothing Then
        ClickButtonByUIA = False
        Exit Function
    End If

    Set oCondition = oAutomation.CreatePropertyCondition(UIA_NamePropertyId, strButtonName)
    Set oButtonElement = oWindowElement.FindFirst(TreeScope_Descendants, oCondition)

    If oButtonElement Is Nothing Then
        ClickButtonByUIA = False
        Exit Function
    End If

    Set oInvokePattern = oButtonElement.GetCurrentPattern(UIA_InvokePatternId)
    oInvokePattern.Invoke

    ClickButtonByUIA = True
    Exit Function

ErrHandler:
    ClickButtonByUIA = False
End Function

' 목적   : EnumChildWindows 콜백 — m_strTargetText 를 포함하는 첫 자식 창을 m_hWndFound 에 기록
Private Function prv_EnumChildProc(ByVal hWndChild As LongPtr, ByVal lParam As LongPtr) As Long
    Dim strBuffer As String
    Dim lngLen    As Long

    strBuffer = String(256, vbNullChar)
    lngLen = GetWindowTextA(hWndChild, strBuffer, 256)
    strBuffer = Left(strBuffer, lngLen)

    If InStr(strBuffer, m_strTargetText) > 0 Then
        m_hWndFound = hWndChild
        prv_EnumChildProc = 0
    Else
        prv_EnumChildProc = 1
    End If
End Function
