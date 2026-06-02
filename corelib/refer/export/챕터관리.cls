VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "챕터관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub SearchBox_Change()
    Dim celTarget As Range
    Set celTarget = [U3]
    With Me.SearchBox
    If .ListIndex = -1 Then
        celTarget = IIf(.value = "", "", "%" & .value & "%")
    Else
        celTarget = .value
    End If
    End With
End Sub
Private Sub SearchBox_DropButtonClick()
    If Me.SearchBox.ListIndex = -1 Then
        Call FormSearchFilter
    End If
End Sub
Private Sub SearchBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then '엔터 키(KeyCode = 13) 입력 시 이벤트 처리
        Call RunShpMacro("SEARCH")
    End If
End Sub

