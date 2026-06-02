VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_AddList 
   ClientHeight    =   4845
   ClientLeft      =   156
   ClientTop       =   516
   ClientWidth     =   3180
   OleObjectBlob   =   "frm_AddList.frx":0000
End
Attribute VB_Name = "frm_AddList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Move Application.Left + (Application.Width - .Width) / 2, Application.Top + (Application.Height - .Height) / 2 ''폼의 위치를 중앙으로 이동
    End With
End Sub
Private Sub bt_addITEM_Click()

    cl.sht_UnLock
    
    Dim arrList()
    Dim strList As String
    Dim i, j As Long
    
    With Me.lst_List
    For i = 0 To .ListCount - 1
        
        If .Selected(i) Then
            ReDim Preserve arrList(j)
            
            arrList(j) = .List(i)
            j = j + 1
            
        End If
        
    Next
    End With
    strList = Join(arrList, ";")
    celDC = strList
    
    cl.sht_Lock
    
    Unload Me
    
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub
