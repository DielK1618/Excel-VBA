Attribute VB_Name = "tpl_KeyBoard"
Option Explicit
' 키보드 이벤트를 위한 Windows API 선언
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
' 키보드 이벤트 상수
Private Const KEYEVENTF_KEYDOWN = 0
Private Const KEYEVENTF_KEYUP = 2
' 자주 사용하는 가상 키코드
Public Enum VirtualKeys
    VK_TAB = &H9         ' Tab
    VK_RETURN = &HD      ' Enter
    VK_CONTROL = &H11    ' Ctrl
    VK_MENU = &H12       ' Alt
    VK_ESCAPE = &H1B     ' Esc
    VK_SPACE = &H20      ' Space
    VK_LEFT = &H25       ' Left Arrow
    VK_UP = &H26         ' Up Arrow
    VK_RIGHT = &H27      ' Right Arrow
    VK_DOWN = &H28       ' Down Arrow
End Enum
' 자주 사용하는 키 동작 정의
Public Enum KeyActions
    ACTION_COPY = 1
    ACTION_PASTE = 2
    ACTION_TAB = 3
    ACTION_ENTER = 4
    ACTION_ALT_TAB = 5
    ACTION_ESCAPE = 6
    ACTION_ARROW_DOWN = 7
    ACTION_ARROW_UP = 8
    ACTION_ARROW_LEFT = 9
    ACTION_ARROW_RIGHT = 10
End Enum
' 키보드 동작 실행 함수
Public Sub ExecuteKeyAction(action As KeyActions, Optional waitAfter As Long = 100)
    Select Case action
        Case ACTION_COPY
            SendKeyCombo VK_CONTROL, 67  ' Ctrl + C
        Case ACTION_PASTE
            SendKeyCombo VK_CONTROL, 86  ' Ctrl + V
        Case ACTION_TAB
            SendKey VK_TAB
        Case ACTION_ENTER
            SendKey VK_RETURN
        Case ACTION_ALT_TAB
            SendKeyCombo VK_MENU, VK_TAB
        Case ACTION_ESCAPE
            SendKey VK_ESCAPE
        Case ACTION_ARROW_DOWN
            SendKey VK_DOWN
        Case ACTION_ARROW_UP
            SendKey VK_UP
        Case ACTION_ARROW_LEFT
            SendKey VK_LEFT
        Case ACTION_ARROW_RIGHT
            SendKey VK_RIGHT
    End Select
    
    If waitAfter > 0 Then Wait waitAfter
End Sub
' 여러 키 동작을 연속으로 실행
Public Sub ExecuteKeySequence(ParamArray actions() As Variant)
    Dim i As Long
    
    For i = LBound(actions) To UBound(actions)
        If IsArray(actions(i)) Then
            ' 배열인 경우 첫 번째 요소는 동작, 두 번째 요소는 대기 시간
            ExecuteKeyAction actions(i)(0), actions(i)(1)
        Else
            ' 단일 값인 경우 기본 대기 시간 사용
            ExecuteKeyAction actions(i)
        End If
    Next i
End Sub
' 셀 범위에 대해 지정된 키 시퀀스 실행
Public Sub ProcessRangeWithKeySequence(sourceRange As Range, ParamArray actions() As Variant)
    Dim cell As Range
    
    For Each cell In sourceRange
        If Not IsEmpty(cell) Then
            cell.Select
            ExecuteKeySequence actions
        End If
    Next cell
End Sub
Private Sub SendKey(KeyCode As Byte, Optional PressAndRelease As Boolean = True)
    keybd_event KeyCode, 0, KEYEVENTF_KEYDOWN, 0
    If PressAndRelease Then
        keybd_event KeyCode, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub
Private Sub SendKeyCombo(ParamArray KeyCodes() As Variant)
    Dim i As Long
    
    For i = LBound(KeyCodes) To UBound(KeyCodes)
        SendKey CByte(KeyCodes(i)), False
    Next i
    
    Wait 50
    
    For i = UBound(KeyCodes) To LBound(KeyCodes) Step -1
        keybd_event CByte(KeyCodes(i)), 0, KEYEVENTF_KEYUP, 0
    Next i
End Sub
Private Sub Wait(milliseconds As Long)
    Application.Wait Now + TimeSerial(0, 0, milliseconds / 1000)
End Sub



