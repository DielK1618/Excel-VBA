VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_RowsAddOrDelete 
   ClientHeight    =   1845
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   2940
   OleObjectBlob   =   "frm_RowsAddOrDelete.frx":0000
End
Attribute VB_Name = "frm_RowsAddOrDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    '폼의 위치를 중앙으로 이동
    Me.StartUpPosition = 0
    Me.Move Application.Left + (Application.Width - Me.Width) / 2, Application.Top + (Application.Height - Me.Height) / 2
    
    Dim strDesc As String
    strDesc = " ※ 행의 개수를 입력" & vbNewLine & " ※ 추가 및 삭제 버튼 클릭" & vbNewLine & " ※ 슷자 미입력 후 삭제시 전체 삭제"
    Me.lbl_Desc = strDesc
End Sub
Private Sub cmd_Add_Click()
        
    Dim tbl As ListObject
    Dim intInput, intRow As Long
    
    intInput = Me.txt_int.value
    
    If intInput <> "" Then
    
        Set tbl = celDC.ListObject
        intRow = celDC.Row - tbl.HeaderRowRange.Row
        
        cl.sht_UnLock
        Call AddTableRows(intInput, intRow, tbl)
        cl.sht_Lock
        
        Unload Me
    
    Else
        MsgBox "행의 개수를 먼저 입력해 주세요!", vbInformation
    End If

End Sub
Private Sub cmd_Delete_Click()

    Dim tbl As ListObject
    Dim intInput, intRow As Long
    
    intInput = Me.txt_int.value
    Set tbl = celDC.ListObject
    
    cl.sht_UnLock
    If intInput <> "" Or intInput = 0 Then

        intRow = celDC.Row - tbl.HeaderRowRange.Row
        Call DelTableRows(intRow, intInput, tbl)

    Else
        Call DelTableAllRows(tbl)
    End If
    cl.sht_Lock
    
    Unload Me
    
End Sub
Private Sub txt_int_Change()
    Dim strInput As String
    strInput = txt_int.value
    
    If strInput <> "" Then
        If Not IsNumeric(strInput) Then
            MsgBox "1이상의 정수만 입력 가능합니다!", vbCritical
            txt_int.value = ""
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub
