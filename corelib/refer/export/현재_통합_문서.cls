VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Workbook_Open()
'
'    cl.DPUpdate_Off
'    cl.Event_Off
'    cl.Calculate_Off
'
'    Dim sht As Worksheet
'    Set sht = ThisWorkbook.Worksheets("Home")
'
'    Call ChangeSheetEvent(sht)
'    Call ReloadDbInfo
'    Call GetFieldMap
'    Call GetFieldMapRev
'    Call ResetTargetPages
'    Call BtSyncStartUpDB
'
'    blnPartEvents = True
'    Set wsPub = sht
'
'    cl.sht_Lock
'    cl.Calculate_On
'    cl.Event_On
'    cl.DPUpdate_On
'
'    Set wsPub = Nothing
'
'End Sub
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    Call SetHotKeys
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim tbl As ListObject
    Dim rng As Range, trng As Range
    Dim cell As Range
    Dim modColIndex As Long

    On Error Resume Next
    Set tbl = ActiveSheet.ListObjects(1)
    If tbl Is Nothing Then Exit Sub

    Set rng = Intersect(Target, tbl.DataBodyRange)
    If rng Is Nothing Then Exit Sub

    Set trng = Range(tbl.ListColumns(2).DataBodyRange, _
                     tbl.ListColumns("MOD").DataBodyRange.Offset(, -1))

    Set rng = Intersect(rng, trng)
    If rng Is Nothing Then Exit Sub

    modColIndex = tbl.ListColumns("MOD").Index

    cl.DPUpdate_Off
    cl.Event_Off
    ActiveSheet.Unprotect Password:=PW

    Dim modColData As Variant
    Dim modColRange As Range
    Set modColRange = tbl.ListColumns("MOD").DataBodyRange
    modColData = modColRange.value  ' [핵심] MOD 열 전체를 배열로 한 번에 읽기

    Dim processedRows As Object
    Set processedRows = CreateObject("Scripting.Dictionary")

    Dim unionRng As Range   ' 실제로 MOD 값을 써야 할 셀들을 누적할 Union 범위
    Dim tblStartRow As Long
    tblStartRow = tbl.DataBodyRange.Row

    For Each cell In rng
        Dim rowKey As Long
        rowKey = cell.Row

        If Not processedRows.Exists(rowKey) Then
            processedRows.Add rowKey, True

            Dim rowRange As Range
            Set rowRange = Intersect(cell.EntireRow, trng)

            If WorksheetFunction.CountA(rowRange) > 0 Then
                Dim arrRow As Long
                arrRow = rowKey - tblStartRow + 1  ' 배열 인덱스 (1부터 시작)

                If modColData(arrRow, 1) = "" Then  ' [핵심] 배열에서 읽어 셀 접근 없이 비교
                    Dim modCell As Range
                    Set modCell = modColRange.Cells(arrRow, 1)

                    If unionRng Is Nothing Then
                        Set unionRng = modCell  ' 첫 번째 대상 셀
                    Else
                        Set unionRng = Union(unionRng, modCell)  ' [핵심] 대상 셀 누적
                    End If
                End If
            End If
        End If
    Next cell

    If Not unionRng Is Nothing Then
        unionRng.value = 1  ' [핵심] Union된 모든 MOD 셀에 한 번에 쓰기
    End If

    Set wsPub = ActiveSheet
    cl.sht_Lock

    cl.Event_On
    cl.DPUpdate_On
End Sub
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    Cancel = True

    Dim tbl As ListObject
    Dim celTblHeader As Range
    Dim strHeader As String
    Dim t As Range
    Dim blnROK As Boolean

    cl.DPUpdate_Off
    cl.Event_Off
    cl.sht_UnLock
    cl.Calculate_Off
    
    blnROK = False
    Set celDC = Target
    
    With Target.Cells(1, 1)
        
    '<< 타겟이 일반 범위일 경우 >> ==================================================================
        Select Case True
        Case .Interior.Color = uiColor.TC05 And .Font.Color = uiColor.TC12 '파일 URL 등록 폼 호출
            MsoDilogType = msoFileDialogFilePicker
            frm_GetPath.Show
        Case .Interior.Color = uiColor.TC05 And .Font.Color = uiColor.TC01 '파일 다이얼로그 호출
            celDC = GetPath(, IIf(celDC.value = "", "", ReplacePath(Mid(celDC.value, 1, InStrRev(celDC.value, "\")))), , , , , celDC.value)
        Case .Interior.Color = uiColor.TC05 And .Font.Color = uiColor.TC02 '폴더 다이얼로그 호출
            celDC = GetPath(msoFileDialogFolderPicker, IIf(celDC.value = "", "", ReplacePath(Mid(celDC.value, 1, InStrRev(celDC.value, "\")))), , , , , celDC.value)
        Case .Interior.Color = uiColor.TC06 And .Font.Color = uiColor.TC03 '연결테이블 폼 호출
            Call AddSelectionItems(GetTwbRange("T_데이터동기화[동기화테이블]"), celDC)
            frm_AddList.Show
        Case .Interior.Color = uiColor.TC06 And .Font.Color = uiColor.TC12 '다중 바이너리 선택
            If .value = "" Then
               Call SetIfValTrue(Target, 1, True)
            Else
               Call SetIfValTrue(Target, "")
            End If
        Case .Interior.Color = uiColor.TC06 And .Font.Color = uiColor.TC02  '텍스트 박스 호출
            frm_TextBox.Show
        End Select
        
    '<< 타겟이 테이블인 경우 >> ==================================================================
    If IsTableRange(Target, tbl, celTblHeader) Then
        
        strHeader = celTblHeader
        
        '입력 범위 지정
        Dim rngBody As Range
        
        On Error Resume Next
        Set rngBody = tbl.ListColumns(strHeader).DataBodyRange
        Set rngBody = rngBody.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not Intersect(Target, celTblHeader) Is Nothing Then '<< 타겟이 테이블의 헤더일 경우 >>
            Select Case True
            Case strHeader = "No" Or (celTblHeader.Column = 1 And strHeader = "ID") '테이블 필터 해지
                Call ClearFiltersInTable(tbl)
            Case rngBody.Interior.Color = uiColor.TC06 And rngBody.Font.Color = uiColor.TC12 '항목 선택
                Dim vntValue As Variant
                vntValue = IIf(Application.WorksheetFunction.Subtotal(3, rngBody) = 0, 1, "")
                Call SetIfValTrue(rngBody, vntValue)
            Case Else
                If InStr(tbl.Name, "TS_") = 0 Then
                    With frm_TableArrayFilter
                        .txt_FieldName.value = strHeader
                        .Show
                    End With
                Else
                    If .Interior.Color = uiColor.TC16 Then
                        If .Font.Color = uiColor.TC12 Then
                            .Font.Color = uiColor.TC13
                        Else
                            .Font.Color = uiColor.TC12
                        End If
                    End If
                End If
            End Select
        Else '<< 타겟이 테이블의 바디일 경우 >>
            Select Case True
            Case .Interior.Color = uiColor.TC06 And .Font.Color = uiColor.TC10 '단일 바이너리 선택
                rngBody.value = ""
                Call SetIfValTrue(Target, 1)
                If tbl.Name = "T_데이터베이스관리" Then Call ReloadDbInfo
            Case strHeader = "No"
                frm_RowsAddOrDelete.Show
            End Select
        End If
    End If
    End With

    cl.Calculate_On
    cl.sht_Lock
    cl.Event_On
    cl.DPUpdate_On

End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
   Sh.Calculate
End Sub
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    If blnPartEvents = False Then Exit Sub
    
    cl.DPUpdate_Off
    cl.Calculate_Off

    '페이지 초기화
    On Error Resume Next
    If TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "테이블리셋", "명칭", "=", Sh.Name)(0) <> "" Then Call BtClearSearchSheet(Sh, False)
    On Error GoTo 0

    Dim rngFind As Range
    Dim f As Range
    
    Set rngFind = Range("1:1")
    Set f = rngFind.Find("업데이트연결", , xlFormulas, xlWhole)
    
    If Not f Is Nothing Then If f.Offset(1) = "MSG ON" Then Call BtUpdateDBonThisPage

    '선택행 강조 조건부 서식 ============================================
    Dim tbl As ListObject
    Dim condFormat As FormatCondition
    Dim formulaToCheck, formulaFinal As String
    Dim i, intFirstCol, intLastCol As Integer

     '조건부 서식의 수식을 설정합니다
    formulaToCheck = "CELL(""ROW"")=ROW()"
    
    cl.sht_UnLock
     '시트의 모든 테이블을 확인합니다
    On Error Resume Next
    For Each tbl In Sh.ListObjects
         '테이블의 DataBodyRange를 설정합니다
        If Not tbl.DataBodyRange Is Nothing Then
             '테이블 범위의 모든 조건부 서식을 확인합니다
            For i = tbl.DataBodyRange.FormatConditions.count To 1 Step -1
                Set condFormat = Nothing
                Set condFormat = tbl.DataBodyRange.FormatConditions(i)
                If condFormat.Type = xlExpression Then
                    If InStr(condFormat.Formula1, formulaToCheck) > 0 Then
                        condFormat.Delete
                    End If
                End If
            Next i
        End If

        If InStr(tbl.Name, "TS") = 0 Then
            intFirstCol = tbl.ListColumns(1).Range.Column
            intLastCol = tbl.ListColumns(tbl.ListColumns.count).Range.Column
            formulaFinal = "=AND(CELL(""COL"")>=" & intFirstCol & ",CELL(""COL"")<=" & intLastCol & "," & formulaToCheck & ")"
            
            Call ConditionalFormattingFormula(tbl.DataBodyRange, formulaFinal, , , xlPatternGray16, uiColor.TC11)
        End If
    Next tbl
    On Error GoTo 0
    
    cl.sht_Lock
    cl.Calculate_On
    cl.DPUpdate_On
    
End Sub


