VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_DBinfo 
   ClientHeight    =   6450
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   4620
   OleObjectBlob   =   "frm_DBinfo.frx":0000
End
Attribute VB_Name = "frm_DBinfo"
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
    
    Dim r, rng As Range
    Set rng = GetTwbRange("T_데이터베이스관리[별칭]")

    With Me.cbo_Aliases
        .Clear
        For Each r In rng
            .AddItem r
        Next
    End With
   
End Sub
Private Sub bt_Refresh_Click()
    Me.cbo_Aliases.ListIndex = -1
    Call BtGetDbTableNames
End Sub
Private Sub cbo_Aliases_Change()
    
    Me.txt_Token = ""
    Me.txt_DB = ""
    
    Dim dv As DBvar
    
    With dv
        .Alias = Me.cbo_Aliases.value
        
        Dim arr As Variant
        arr = GetDbInfo(.Alias, False)
        .Type = arr(0)
        .Token = arr(2)
        .db = arr(4)
        
        Me.txt_Type = .Type
        If Me.txt_Token <> .Token Then Me.txt_Token = .Token
        Me.txt_DB = .db
        
        Dim r, rng As Range
        Dim strKey As String
        
        Me.cbo_Tables.Clear
        Me.lst_Fields.Clear
        
        If .Token = "TWB_DB" Then
            Dim ws As Worksheet
            
            For Each ws In ThisWorkbook.Worksheets
                If InStr(ws.Name, "DB_") > 0 Then
                  Me.cbo_Tables.AddItem ws.Name
                End If
            Next
        Else
            Set rng = GetTwbRange("T_외부테이블[테이블]")
            
            For Each r In rng
                strKey = .Token & .db
                
                If (r.Offset(, intOffset(r, "예약어")) & r.Offset(, intOffset(r, "데이터베이스"))) = strKey Then
                    Me.cbo_Tables.AddItem r
                End If
            Next
        End If
        
    End With
    
End Sub
Private Sub txt_Token_Change()
    
    If Me.cbo_Aliases = "" Then
    
    Dim dv As DBvar
    
    With dv
        Dim arr As Variant
        arr = GetDbInfo(.Token)
        .Alias = arr(1)
        Me.cbo_Aliases = .Alias
    End With
    
    End If
    
End Sub
Private Sub cbo_Tables_Change()
    
    Dim arr As Variant
    Dim strToken, strTable As String

    With Me
        strToken = .txt_Token
        strTable = .cbo_Tables
        
        If strToken = "" Or strTable = "" Then Exit Sub
        
        arr = GetFieldNameConnection(strToken, strTable)
    End With
    With Me.lst_Fields
        .ColumnCount = 3
        .Clear
        .List = Application.WorksheetFunction.Transpose(arr)
    End With
        
End Sub
Private Sub bt_Input_Click()
    Dim cel As Range
    Set cel = Selection
    
    On Error Resume Next
    If IsCells(cel) Then
        With Me
            Select Case True
            Case .opt_Aliases
                cel = .cbo_Aliases
            Case .opt_Type
                cel = .txt_Type
            Case .opt_Token
                cel = .txt_Token
            Case .opt_DB
                cel = .txt_DB
            Case .opt_Tables
                cel = .cbo_Tables
            Case .opt_Fields
                cel = .lst_Fields.List(.lst_Fields.ListIndex, 0)
            Case Else
                MsgBox "입력할 항목을 선택하세요!", vbInformation
            End Select
        End With
    End If
    On Error GoTo 0
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub
