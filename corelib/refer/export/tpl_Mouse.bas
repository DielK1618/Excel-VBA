Attribute VB_Name = "tpl_Mouse"
Option Explicit
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' POINTAPI 구조체 선언
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

' 마우스 이벤트 상수 정의
Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Const MOUSEEVENTF_LEFTUP As Long = &H4
Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Const MOUSEEVENTF_RIGHTUP As Long = &H10
Function GetMousePosition(ByRef x As Variant, ByRef y As Variant)  '마우스 포인트의 위치 확인
    Dim mousePos As POINTAPI
    
    ' 현재 마우스 좌표를 가져옴
    GetCursorPos mousePos
    ' 좌표 출력
    x = mousePos.x
    y = mousePos.y
End Function
Sub ClickAtPosition(x As Long, y As Long, Optional BoolLeft As String = True, Optional strDelay As String, Optional strWait As String) '마우스를 입력한 좌표로 이동 후 클릭
    ' 마우스 커서를 지정된 위치로 이동
    SetCursorPos x, y
    
    If strDelay <> "" Then WaitTime strDelay
    
    ' 버튼 타입에 따라 다른 이벤트 실행
    Select Case BoolLeft
        Case True
            mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Case False
            mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
            mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    End Select
    
    If strWait <> "" Then WaitTime strWait
    
End Sub
Sub GetMousePosition_BT()

    Dim x, y As Variant
    
    Call GetMousePosition(x, y)
    
    Debug.Print x
    Debug.Print y
    
End Sub
Sub WaitTime(ByVal strTime As String) '입력한 시간(hh:mm:ss)만큼 대기
    Application.Wait (Now + TimeValue(strTime))
End Sub

