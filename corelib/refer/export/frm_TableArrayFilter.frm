VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_TableArrayFilter 
   ClientHeight    =   7815
   ClientLeft      =   156
   ClientTop       =   612
   ClientWidth     =   3684
   OleObjectBlob   =   "frm_TableArrayFilter.frx":0000
End
Attribute VB_Name = "frm_TableArrayFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    '폼의 위치를 중앙으로 이동
    Me.Move Application.Left + (Application.Width - Me.Width) / 2, Application.Top + (Application.Height - Me.Height) / 2
End Sub
Private Sub bt_Clear_Click()
    Dim tbl As ListObject
    Set tbl = celDC.ListObject
    Call ClearFiltersInTable(tbl)
    
    Unload Me
End Sub
Private Sub bt_Apply_Click()

    Dim tbl As ListObject
    Set tbl = celDC.ListObject
    
    If Me.opt_New.value = True Then
        Call ClearFiltersInTable(tbl)
    End If
    
    Dim arrValues() As String
    Dim strValues As String
    Dim i As Long
    
    arrValues = Split(Me.txt_Values.value, vbNewLine)
    
    Call AutoTableFilter_Arr(tbl, Me.txt_FieldName.value, arrValues, Me.tgl_Part.value)
    
    Unload Me
    
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub

