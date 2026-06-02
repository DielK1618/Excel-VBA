VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_TextBox 
   Caption         =   "텍스트박스"
   ClientHeight    =   5415
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   4620
   OleObjectBlob   =   "frm_TextBox.frx":0000
End
Attribute VB_Name = "frm_TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Move Application.Left + (Application.Width - .Width) / 2, Application.Top + (Application.Height - .Height) / 2 ''폼의 위치를 중앙으로 이동
    
    If celDC Is Nothing Then
        .bt_Save.Enabled = False
        .tgl_Formula.Visible = False
        GoTo pass
    End If
    If celDC(1).HasFormula Then
        .bt_Save.Enabled = False
        .tgl_Formula.Visible = True
    Else
        .bt_Save.Enabled = True
        .tgl_Formula.Visible = False
    End If
    '더블클릭된 셀의 내용을 텍스트 박스에 출력
    .txt_TextBox = celDC(1)
pass:
    ' 스크롤을 맨 위로 이동
    .txt_TextBox.SelStart = 0
    .txt_TextBox.SelLength = 0
    
    End With
End Sub
Private Sub bt_Save_Click()
    cl.sht_UnLock
    Dim strText As String
    strText = Me.txt_TextBox.value
    
    ' 빈 값 처리
    If Trim(strText) = "" Then
        celDC.value = ""
        GoTo ExitSub
    End If
    
    ' 수식 여부 판별
    If Left(LTrim(strText), 1) = "=" Then
        ' ── 수식 입력 ──
        celDC.Formula = strText
    Else
        ' ── 텍스트 입력 (줄바꿈 유지) ──
        strText = Replace(strText, vbCrLf, Chr(10))
        strText = Replace(strText, vbLf, Chr(10))
        
        celDC.value = strText
        celDC.WrapText = True
    End If
    
ExitSub:
    cl.sht_Lock
    Unload Me
End Sub
Private Sub tgl_Formula_Click()
    
    With Me
        If .tgl_Formula.value Then
            .txt_TextBox.value = celDC.Formula
            .bt_Save.Enabled = True
        Else
            .txt_TextBox.value = celDC
            .bt_Save.Enabled = False
        End If
    End With
    
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub


