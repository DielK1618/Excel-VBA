Attribute VB_Name = "tpl_Sheet"
Option Explicit
Sub HideAllSheetsExceptOne(ByVal sht As Worksheet)
            
    On Error Resume Next
    Dim ws As Object
    
    With sht
        .Visible = True
        .Activate
        .Range("A1").Activate '활성화할 범위
    
    Dim wb As Workbook

    For Each ws In ThisWorkbook.Sheets ' 모든 시트 숨기기
        If ws.Name <> .Name Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    
    End With
End Sub
Sub ChangeSheetEvent(ByVal shtAct As Worksheet)
    
    ThisWorkbook.Activate
    
    Dim shtDeact As Worksheet
    Set shtDeact = ActiveSheet
    Set wsPub = shtDeact
    
    cl.WB_UNLock
    cl.sht_UnLock
    
    If blnPartEvents Then
        '비활성화 될 시트 설정 ==========================

        Dim tblSet As ListObject
        Set tblSet = GetTwbRange("T_페이지설정").ListObject
        
        Dim vntReset As Variant
        vntReset = TblFindVals_MC(tblSet, "테이블리셋", "명칭", "=", shtDeact.Name)
        
        Dim tbl As ListObject
        
        '테이블을 순회하며 필터 해지 및 데이터 삭제
        On Error Resume Next
        For Each tbl In shtDeact.ListObjects
            Call ClearFiltersInTable(tbl)
            If vntReset(0) = 1 And InStr(tbl.Name, "DB_") = 0 Then Call DelTableAllRows(tbl)
        Next
        
        '스크롤 초기화
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        On Error GoTo 0
    End If
    
    cl.sht_Lock
    
    '활성화 될 시트 설정 ==========================
    Call HideAllSheetsExceptOne(shtAct)
    
    Set wsPub = shtAct
    cl.sht_UnLock
    
    If blnPartEvents Then
        vntReset = TblFindVals_MC(tblSet, "테이블리셋", "명칭", "=", shtAct.Name)
        
        '테이블을 순회하며 필터 해지 및 데이터 삭제
        On Error Resume Next
        For Each tbl In shtAct.ListObjects
            Call ClearFiltersInTable(tbl)
            If vntReset(0) = 1 And InStr(tbl.Name, "DB_") = 0 Then Call DelTableAllRows(tbl)
        Next
        
        '스크롤 초기화
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        On Error GoTo 0
    End If
    
    Call SearchBoxResize
    
    cl.sht_Lock
    cl.WB_Lock
    
End Sub
Sub VisibleAllSheets()
    
    On Error Resume Next
    Dim ws As Object
    
    For Each ws In Sheets
            ws.Visible = xlSheetVisible
    Next ws

End Sub
Function GetSheetNames() As Variant
    
    Dim arrResult()
    Dim sht As Object
    Dim i As Long
    
    ReDim arrResult(0)
    arrResult(0) = ""
    
    For Each sht In ThisWorkbook.Sheets
        ReDim Preserve arrResult(i)
        arrResult(i) = sht.Name
        i = i + 1
    Next
    
    GetSheetNames = Application.WorksheetFunction.Transpose(arrResult)
    
End Function

Sub SortSheets(ByVal arrSheetNames As Variant) '메뉴관리
    
    Call VisibleAllSheets
    
    Dim shtMove, shtDestination As Object
    Dim shtName As Variant
    
    For Each shtName In arrSheetNames
        On Error Resume Next
        Set shtMove = ThisWorkbook.Sheets(CStr(shtName))
        
            If Not shtMove Is Nothing Then
                shtMove.Move After:=ThisWorkbook.Sheets(Sheets.count)
                'shtMove.Tab.Color = uiColor.Color01
            End If
    Next
    
End Sub
