VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Navigation 
   ClientHeight    =   11550
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   4812
   OleObjectBlob   =   "frm_Navigation.frx":0000
End
Attribute VB_Name = "frm_Navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Left = Application.Left + Application.Width - .Width - 30
        .Top = Application.Top + Application.Height - .Height - 50
        .Tab_Navigation.Tabs.Clear
        '대분류 메뉴 삽입
        Dim r, rng As Range
        Set rng = GetTwbRange("T_분류설정[명칭]")
    
        For Each r In rng
            If r.Offset(, intOffset(r, "보기")) = 1 Then .Tab_Navigation.Tabs.Add r
        Next r
    End With
End Sub
Private Sub Tab_Navigation_Change()
    
    ThisWorkbook.Activate
        
    Dim r, rng As Range
    Dim s As String

    Set rng = GetTwbRange("T_페이지설정[명칭]")
    s = Tab_Navigation.SelectedItem.Name
    
    lst_Navigation.Clear
    lst_ListMemo.Caption = ""
    
    For Each r In rng
        If r.Offset(, intOffset(r, "분류")) = s Then
            If r.Offset(, intOffset(r, "보기")) = 1 Then
                lst_Navigation.AddItem r
            End If
        End If
    Next
End Sub
Private Sub lst_Navigation_Click()
    
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(lst_Navigation.List(lst_Navigation.value))
    
    cl.DPUpdate_Off
    Call ChangeSheetEvent(sht)
    cl.DPUpdate_On
    
    lst_ListMemo.Caption = CStr(TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "설명", "명칭", "=", sht.Name)(0))
    
End Sub
Private Sub bt_Home_Click()

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Home")
    
    cl.DPUpdate_Off
    Call ChangeSheetEvent(sht)
    cl.DPUpdate_On
    
   lst_ListMemo.Caption = CStr(TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "설명", "명칭", "=", sht.Name)(0))
   
End Sub
Private Sub bt_Group_Click()

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("분류설정")
    
    cl.DPUpdate_Off
    Call ChangeSheetEvent(sht)
    cl.DPUpdate_On
    
    lst_ListMemo.Caption = CStr(TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "설명", "명칭", "=", sht.Name)(0))
    
End Sub
Private Sub bt_Setting_Click()

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("페이지설정")
    
    cl.DPUpdate_Off
    Call ChangeSheetEvent(sht)
    cl.DPUpdate_On
    
    lst_ListMemo.Caption = CStr(TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "설명", "명칭", "=", sht.Name)(0))
    
End Sub
