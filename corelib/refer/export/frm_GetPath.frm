VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_GetPath 
   Caption         =   "Get Path"
   ClientHeight    =   1230
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   5412
   OleObjectBlob   =   "frm_GetPath.frx":0000
End
Attribute VB_Name = "frm_GetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Move Application.Left + (Application.Width - .Width) / 2, Application.Top + (Application.Height - .Height) / 2 ''폼의 위치를 중앙으로 이동
        .txt_url.value = celDC.value '텍스트 박스의 초기 값
    End With
End Sub
Private Sub bt_Selector_Click()
    Dim strExistingPath As String
    
    With Me.txt_url
    strExistingPath = .value
    .value = GetPath(MsoDilogType, IIf(strExistingPath = "", "", ReplacePath(Mid(strExistingPath, 1, InStrRev(strExistingPath, "\")))), , , , vntTemp, strExistingPath)
    End With
    
End Sub
Private Sub bt_Save_Click()
    Dim strPath As String
    strPath = ReplacePath(Me.txt_url)

    If CheckFileExistence(strPath) = False Then If CheckFolderExistence(strPath) = False Then GoTo PathErr
    celDC.value = Trim(Me.txt_url)
    Unload Me
    Exit Sub
PathErr:
    If MsgBox("파일 경로가 유효하지 않습니다. 그래도 등록하시겠습니까?", vbCritical + vbYesNo) = vbYes Then
        celDC.value = Me.txt_url
        Unload Me
    End If
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub
