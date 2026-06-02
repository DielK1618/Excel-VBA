VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_UnProtect 
   ClientHeight    =   1560
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4704
   OleObjectBlob   =   "frm_UnProtect.frx":0000
End
Attribute VB_Name = "frm_UnProtect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Move Application.Left + (Application.Width - .Width) / 2, Application.Top + (Application.Height - .Height) / 2 ''폼의 위치를 중앙으로 이동
                
        With .txt_Password
            .Text = ""
            .PasswordChar = "*"
            .IMEMode = fmIMEModeOff   ' 영어 모드
        End With
    End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' X 버튼으로 닫을 때도 Tag 초기화 후 Hide (Unload 방지)
    If CloseMode = 0 Then
        Cancel = True
        Me.Tag = ""
        Me.Hide
    End If
End Sub
' Enter 키로도 Unlock 실행
Private Sub txt_Password_Enter()
    Me.txt_Password.IMEMode = fmIMEModeOff   ' 영어 입력 모드
End Sub
'Private Sub txt_Password_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
'                                  ByVal Shift As Integer)
'    If KeyCode = 13 Then bt_Unlock_Click
'End Sub
Private Sub bt_Continue_Click()
    With Me
        .Tag = txt_Password.value   ' Tag에 암호 저장
        
        If .Tag = "" Then
            MsgBox "[비밀번호]를 입력하세요!", vbInformation
            .txt_Password.SetFocus
            Exit Sub
        End If
        .Hide
    End With
End Sub
Private Sub bt_Cancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub


