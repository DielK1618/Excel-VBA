Attribute VB_Name = "tpl_Buttons_Top"
Option Explicit
Sub BtNavigation() '네비게이션 메뉴 호출
    frm_Navigation.Show 0
End Sub
Sub BtBakupWorkBook() '워크북 백업

    If MsgBox("[현재 문서] 및 [데이터베이스] 파일을 백업하시겠습니까?", vbQuestion + vbYesNo, "파일백업") = vbYes Then
        
        cl.DPUpdate_Off
        
        Dim Path As String
        Path = ReplacePath("{tPath}\00_Bak\")
        
        If CheckFolderExistence(Path) = False Then Call MkFolder(Path)
        
        Call BackupWorkbook(Path)
        Call BackUpDB
        MsgBox "[원본경로\00_Bak\]에 백업 되었습니다.", vbInformation
        
        cl.DPUpdate_On
    End If
    
End Sub
Sub BackUpDB()
    
    Dim r As Range, rng As Range
    Set rng = GetTwbRange("T_데이터베이스관리[파일]")
    
    Dim arr As Variant
    Dim strTwb As String, strPath As String, strRen As String
    
    strTwb = ReplacePath("{tFile}")
    strTwb = LCase(Mid(strTwb, InStrRev(strTwb, "\") + 1))
    
    For Each r In rng
        If r <> "" And InStr(LCase(r), strTwb) = 0 Then
            strPath = ReplacePath(r)
            strRen = Mid(strPath, 1, InStrRev(strPath, "\")) & "00_bak\" & "Bak_" & Format(Now, "(yymmdd_hhmmss) ") & Mid(strPath, InStrRev(strPath, "\") + 1)
            
            Call CopyFile(strPath, strRen)
        End If
    Next
End Sub
Sub BtSetHotKeys() '단축키 설정
    Call SetHotKeys(True)
End Sub
Sub BtSyncStartUpDB()
    Call BtSyncDB(, True, False)
End Sub
Sub BtGetDbTableNames() '모든 데이터 테이블 업데이트
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    Set wsPub = ThisWorkbook.Worksheets("테이블관리")
    
    cl.DPUpdate_Off
    cl.Event_Off
    cl.Calculate_Off
    cl.sht_UnLock
    
    Dim arrClass As Variant
    
    arrClass = Array("시트", "테이블", "외부테이블")
    
    Dim tbl As ListObject
    Dim arrResult As Variant
    Dim celTarget As Range
    Dim strTableName As String
    Dim i As Long
    
    For i = 0 To UBound(arrClass)
    
        Set tbl = GetTwbRange("T_" & arrClass(i)).ListObject
        Call DelTableAllRows(tbl)
        Set celTarget = tbl.ListColumns(2).Range(2, 1)
        
        Select Case arrClass(i)
        Case "시트"
            arrResult = GetSheetNames
            celTarget.Resize(UBound(arrResult), UBound(arrResult, 2)) = arrResult
            Call SortTable(tbl.Name, "명칭")
        Case "테이블"
            arrResult = GetAllSheetTableNames
            celTarget.Resize(UBound(arrResult), UBound(arrResult, 2)) = arrResult
            Call SortTable(tbl.Name, "시트")
        Case "외부테이블"
            Call GetAllDbTableNames(celTarget)
        End Select
        
        Erase arrResult
    Next
    
    MsgBox "Completion!", vbInformation
    
    cl.sht_Lock
    cl.Calculate_On
    cl.Event_On
    cl.DPUpdate_On
    
    Set wsPub = Nothing
    
End Sub
Sub BtSortSheets()
    
    cl.DPUpdate_Off
    cl.Event_Off
    
    Dim arrSheetNames()
    Dim r, rng As Range
    Dim shtActive As Worksheet
    Dim i As Long
    
    Set shtActive = ActiveSheet
    Set rng = GetTwbRange("T_페이지설정[명칭]")
    
    
    For Each r In rng
        ReDim Preserve arrSheetNames(i)
        arrSheetNames(i) = r
        i = i + 1
    Next

    Call SortSheets(arrSheetNames)
    Call HideAllSheetsExceptOne(shtActive)
    
    MsgBox "Completion!", vbInformation
    
    cl.Event_On
    cl.DPUpdate_On
    
End Sub
Sub BtCollectiveReg()
    
    cl.DPUpdate_Off
    cl.sht_UnLock
    cl.Calculate_Off
    
    Dim shp As Shape
    Dim tbl, tblTarget As ListObject
    Dim h, rngTarget As Range
    
    ' 도형 객체 참조
    Set shp = ActiveSheet.Shapes(Application.Caller)
    Set tbl = Range(shp.TopLeftCell.Address).Offset(, -1).ListObject
    Set tblTarget = ActiveSheet.ListObjects(1) 'Range(shp.TopLeftCell.Address).Offset(4).ListObject
    
    On Error Resume Next
    For Each h In tbl.HeaderRowRange
        If h.Interior.Color = uiColor.TC16 And h.Font.Color = uiColor.TC12 Then
        
            Set rngTarget = Range(tblTarget.Name & "[" & CStr(h) & "]")
            If rngTarget.Rows.count > 1 Then Set rngTarget = rngTarget.SpecialCells(xlCellTypeVisible)
            
            rngTarget = tbl.ListColumns(CStr(h)).DataBodyRange(1, 1)
        
        End If
    Next
    On Error GoTo 0
    
    cl.Calculate_On
    cl.sht_Lock
    cl.DPUpdate_On
    
End Sub
Sub BtSyncDB(Optional ByVal arrTables As Variant, _
                            Optional ByVal blnStartUp As Boolean = False, _
                            Optional ByVal blnMsgBox As Boolean = True)

    If blnMsgBox Then
        If MsgBox("데이터베이스 업데이트를 실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        cl.DPUpdate_Off
        cl.Calculate_Off
    End If
    
    Dim tbl As ListObject
    Dim r, rng As Range
    Dim strTarget As String
    Dim intStartUp As Integer
    
    Set tbl = GetTwbRange("T_데이터동기화[선택]").ListObject
    
    If IsArrayEmpty(arrTables) Then 'arrTables 배열에 값이 없는 경우
    
        If blnStartUp = False Then '선택된 항목만 대상 범위로 지정
            Set rng = tbl.ListColumns("선택").DataBodyRange
        Else
            Set rng = tbl.ListColumns("시작실행").DataBodyRange
        End If
        
        Dim intCount As Long
        intCount = Application.WorksheetFunction.CountIf(rng, 1)
        
        If intCount > 0 Then
            Set wsPub = rng.Worksheet
            cl.sht_UnLock
            Set rng = rng.SpecialCells(xlCellTypeConstants, xlNumbers)
            cl.sht_Lock
            Set wsPub = Nothing
        Else
            Exit Sub
        End If

    Else 'arrTables 배열에 값이 있는 경우
    
        Dim arr() As String
        Dim strConditions As String
        Dim i As Long
        
        For i = 0 To UBound(arrTables)
            strConditions = strConditions & IIf(i = 0, "", ";OR;") & "동기화테이블;=;" & arrTables(i)
        Next
        
        arr = Split(strConditions, ";")
        Set rng = TblFindRng_MC(tbl, "동기화테이블", arr)
    End If
        
    For Each r In rng
        strTarget = r.Offset(, intOffset(r, "동기화테이블"))
        If strTarget <> "DB_업데이트로그" Then Call SyncDB(strTarget)
    Next r
    
    Call SyncDB("DB_업데이트로그")
    
     If blnMsgBox Then
        MsgBox "Completion!", vbInformation
        cl.Calculate_On
        cl.DPUpdate_On
    End If
    
End Sub
Sub BtUpdateDBonThisPage(Optional ByVal blnMsgBox As Boolean = True)

    '관계 데이터베이스 업데이트
    Dim rngFind As Range
    Dim f As Range
    Dim strTable As String
    
    Set rngFind = Range("1:1")
    Set f = rngFind.Find("업데이트연결", , xlFormulas, xlWhole)
    
    If Not f Is Nothing Then

        strTable = f.Offset(2)
        If strTable <> "" Then
        
            Dim arrTables() As String
            Dim i As Integer
            
            arrTables = Split(strTable, ";")
            strTable = ""
            
            For i = 0 To UBound(arrTables)
                strTable = strTable & IIf(strTable = "", "", vbNewLine) & "   " & i + 1 & ". " & arrTables(i)
            Next
            
            If blnMsgBox Then If MsgBox("이 페이지의 관계 테이블을 업데이트하시겠습니까?" & vbNewLine & vbNewLine & _
                                                                "※ 관계 테이블 목록" & vbNewLine & _
                                                                strTable, vbYesNo + vbQuestion) <> vbYes Then Exit Sub
            
            
            Call BtSyncDB(arrTables, , False)
            
            If blnMsgBox Then MsgBox "Completion!", vbInformation
        
            End If
    End If
    
End Sub
Sub SyncDB(ByVal strTarget As String)

    Dim tblFind As ListObject
    Dim f As Range, rng As Range
    Dim strAlias  As String
    Dim strTable As String
    Dim strQPlus As String
    Dim strCstQuery As String
    
    On Error GoTo ErrorHandler
    Set tblFind = GetTwbRange("T_데이터동기화").ListObject
    Set rng = tblFind.ListColumns("동기화테이블").DataBodyRange
    
    Set f = rng.Find(strTarget, , xlValues, xlWhole)
    
    If Not f Is Nothing Then
        strAlias = f.Offset(, intOffset(f, "데이터베이스"))
        strTable = f.Offset(, intOffset(f, "테이블"))
        strQPlus = f.Offset(, intOffset(f, "쿼리+"))
        strCstQuery = f.Offset(, intOffset(f, "커스텀쿼리"))
        
        If strAlias = "" Or (strCstQuery = "" And strTable = "") Then GoTo ErrorHandler
    Else
        GoTo ErrorHandler
    End If
    
    Dim dv As DBvar
    With dv
    
        Set .tbl = GetTwbRange(strTarget).ListObject
        .Alias = strAlias
        
        '1. 데이터베이스 ============
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)
        .Server = arrDB(5)
        .Port = arrDB(6)
        .ID = arrDB(7)
        .PW = arrDB(8)
        
        '2. 쿼리 ================
        .cstQuery = strCstQuery
                
        If .cstQuery <> "" Then '커스텀 쿼리 적용
            .Query = .cstQuery
        Else '자동 생성 쿼리 적용
            
            .Table = strTable
            
            If .Table = "" Then
                MsgBox "데이터 테이블을 입력하세요!", vbCritical
                Exit Sub
            End If
            
            .Table = IIf(.Type = "엑셀", "[" & .Table & "$]", .Table)
            
            '필드
            Dim boolParens As Boolean
            If .Type = "엑세스" Then boolParens = True
            
            .arrFields = ReplaceFields(GetFields(, , .tbl), .Token, .Table, boolParens)
            .fields = Join(.arrFields, ", ")
            
            .QPlus = strQPlus
            .Query = "SELECT " & .fields & " FROM " & .Table & " " & .QPlus
            
        End If
        
        .arrQuery = Array(.Query)
        
        '3. 출력범위 ================
        Set .Target = .tbl.Range(2, 1)
        
        '4. 테이블 초기화 ================
        
        Call ClearFiltersInTable(.tbl)
        Call DelTableAllRows(.tbl)
        
        '5. 데이터 로드 ================
        Call RecDBUpdateDateTime(.tbl.Name)
        Call SelectQuery(.Target, .arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
        
    End With
    ' 에러 발생 위치와 상세 정보 출력
    Exit Sub
ErrorHandler:
    Debug.Print "[" & strTarget & "] 테이블에 대한 동기화 정보를 확인하세요!"
End Sub
Sub BtWindowScroll()
    
    Dim arrInt() As String

    arrInt = Split(Application.Caller, "_")

    ActiveWindow.ScrollColumn = arrInt(0)
    ActiveWindow.ScrollRow = arrInt(1)
    
End Sub
Sub BtClearSearchSheet(Optional ByVal sht As Worksheet, _
                                            Optional ByVal blnEvent As Boolean = True)
    
    If sht Is Nothing Then Set sht = ActiveSheet
    Set wsPub = sht
    
    cl.sht_UnLock
    
    If blnEvent Then
        cl.DPUpdate_Off
        cl.Event_Off
        cl.Calculate_Off
    End If
    
    Dim shtTemp As Worksheet
    Set shtTemp = ActiveSheet
    
    If shtTemp.Name = sht.Name Then
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    
    On Error Resume Next
    Call ClearAllComboBoxes(sht)
    
    Dim tbl As ListObject
    
    If sht.ListObjects.count > 0 Then
        For Each tbl In sht.ListObjects
            Call ClearFiltersInTable(tbl)
            Call DelTableAllRows(tbl)
        Next
    Else
        Dim celTarget As Range
        Set celTarget = sht.Range("B9")
        
        If celTarget <> "" Then
            celTarget.CurrentRegion.AutoFilter = False
            Call ClearContents(celTarget)
        End If
    End If

    On Error GoTo 0
    
    If blnEvent Then
        cl.Calculate_On
        cl.Event_On
        cl.DPUpdate_On
    End If
    
    cl.sht_Lock
    Set wsPub = Nothing
End Sub
Sub TestCursorPosition()
    If MsgBox("실행하면 3초 후에 테스트가 시작되니 마우스를 만지지 마세요!", vbInformation + vbYesNo) <> vbYes Then Exit Sub
    
    ClickAtPosition 110, 195, False, "00:00:03", "00:00:01" '친구등록 메뉴 호출
    ClickAtPosition 170, 310, , "00:00:01", "00:00:01" '친구등록 버튼 클릭
    
    MsgBox "Completion!", vbInformation
End Sub
Sub RegisterIDonMessenger_BT() '교육 미이수자 포함 당회 당회장의 메신저 아이디를 쪽지에 등록
    
    If MsgBox("실행하시겠습니까?" & vbNewLine & "실행 후에는 마우스와 키보드 사용을 자제해 주시길 바랍니다!", vbQuestion + vbYesNo, "파일백업") <> vbYes Then Exit Sub
    
    WaitTime "00:00:03"
    
    cl.DPUpdate_Off
    cl.Event_Off
    cl.sht_UnLock
    
    Dim r, rng As Range
    Dim intColTarget As Integer
    Dim i As Long
    
    Set rng = Range("T_당회별수강현황[코드]").SpecialCells(xlCellTypeVisible)
    intColTarget = intOffset(rng, "안내대상")
    blnStop = False
    
    
    frm_Stop.Show 0
    
    For Each r In rng
        If r.Offset(, intColTarget) = 1 Then
        
            DoEvents
            If blnStop Then GoTo pass
            If blnStop = False Then r.Copy
            If blnStop = False Then ClickAtPosition 110, 195, False, "00:00:01", "00:00:01" '친구등록 메뉴 호출
            If blnStop = False Then ClickAtPosition 170, 310, , "00:00:01", "00:00:01" '친구등록 버튼 클릭
            If blnStop = False Then SendKeys "^v", True     '아이디 붙여 넣기
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{TAB}", True '검색 버튼으로 이동
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{ENTER}", True '검색
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{TAB}", True '검색 버튼으로 다시 이동
            If blnStop = False Then SendKeys "{TAB}", True '추가 버튼으로 이동
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{ENTER}", True '추가
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{ENTER}", True '추가 확인
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{ENTER}", True '완료 확인
            If blnStop = False Then WaitTime "00:00:01"
            If blnStop = False Then SendKeys "{TAB}", True '닫기 버튼으로 이동
            If blnStop = False Then SendKeys "{ENTER}", True '닫기
            
            Application.CutCopyMode = False
            i = i + 1
        End If
    Next
    
    AppActivate Application.Caption  '엑셀로 포커스 이동
    Unload frm_Stop
    
    MsgBox "Completion!", vbInformation
pass:
    
    cl.sht_Lock
    cl.Event_On
    cl.DPUpdate_On
    
End Sub
Sub BtGetMonthlyReport()
    
    Dim r As Range, rng As Range
    
    Set rng = GetTwbRange("C4:E4")
    
    If Application.WorksheetFunction.CountBlank(rng) <> 0 Then
        MsgBox "[Year]과 [Month] 값을 모두 입력해 주세요!", vbInformation
        Exit Sub
    End If
    
    cl.DPUpdate_Off
    cl.Calculate_Off
    cl.Event_Off
    cl.sht_UnLock
    
    '검색 연, 월 복사 ====================================
    GetTwbRange("I2") = GetTwbRange("C4")
    GetTwbRange("I3") = GetTwbRange("E4")
    
    
    '주별 이수 참여 현황 데이터 범위 리셋 ====================================
    Dim intCol As Integer
        
    intCol = Application.WorksheetFunction.CountIf(Range("P29:T29"), "<>")
    Set rng = Range("O31")
    Set rng = rng.Resize(4, intCol)
    
    Call SetChartDataRange("CT_주별참여현황", rng)
    
    
    '개인별 이수 현황 테이블 리셋 ====================================
    Set rng = Range("K2:S2")
    
    Dim dv As DBvar
    Dim arrDB As Variant
    Dim arr As Variant

    With dv
        .Alias = "현재워크북"

        arrDB = GetDbInfo(.Alias, False)
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
    
        '데이터를 표에 삽입
        For Each r In rng
            On Error Resume Next
            .Query = r.Offset(1) '변수에 쿼리 삽입

            If .Query = "" Then GoTo pass '쿼리가 비어 있으면 패스

            .arrQuery = Array(.Query) '쿼리를 배열에 담음

            Set .tbl = GetTwbRange(r).ListObject
            Set .Target = .tbl.ListColumns(2).Range(2, 1)

            Call ClearFiltersInTable(.tbl)
            Call DelTableAllRows(.tbl)

            arr = SelectQueryArr(.arrQuery, .Type, .File, , , , , , , True)
            If IsArrayEmpty(arr) Then GoTo pass

            .arrData = arr(0)

            Call AddTableRows(UBound(.arrData, 1), , .tbl)
           .Target.Resize(UBound(.arrData, 1) + 1, UBound(.arrData, 2) + 1).value = .arrData
pass:
            .Query = ""
            If IsArray(arr) Then Erase arr
            If IsArray(.arrData) Then Erase .arrData
            Set .tbl = Nothing
            Set .Target = Nothing
        On Error GoTo 0
        Next r

        cl.Calculate_On
        
        Call WaitTime("00:00:02")
        
        For Each r In rng
            Set .tbl = GetTwbRange(r).ListObject
            Call SortTable(.tbl, "누적", xlDescending)
            Call SortTable(.tbl, "전월", xlDescending)
            Call SortTable(.tbl, "당월", xlDescending)
        Next
    End With
    
    MsgBox "Completion!", vbInformation

    cl.sht_Lock
    cl.Event_On
    cl.DPUpdate_On

End Sub
Sub BtResetMonthlyReportPage()

    cl.DPUpdate_Off
    cl.Calculate_Off
    cl.Event_Off
    cl.sht_UnLock
    
    '검색값 삭제 ======================
    GetTwbRange("I2:I3,C4,E4").ClearContents
    
    '개인별 이수현황 테이블 초기화
    Dim tbl As ListObject
    
    For Each tbl In ThisWorkbook.Worksheets("월간보고").ListObjects
        If InStr(tbl.Name, "타이틀") = 0 Then
            Call ClearFiltersInTable(tbl)
            Call DelTableAllRows(tbl)
        End If
    Next
    
    cl.sht_Lock
    cl.Event_On
    cl.Calculate_On
    cl.DPUpdate_On
End Sub
Sub BtSetMonthlyReportPrintPage()
    
    cl.DPUpdate_Off
    
    Dim ws As Worksheet
    Dim rng As Range, rngTitle As Range
    Set ws = ThisWorkbook.Worksheets("월간보고")
    
    '첫페이지 설정
    Set rng = GetTwbRange("C15:L54")
    Set rngTitle = GetTwbRange("59:59")
    Call SetPrintPage(ws, rng)
    
    '개인별 이수 현황 설정
    Set rng = ws.Range("C56", Cells(Rows.count, 12).End(xlUp))
    Call SetPrintPage(ws, rng, rngTitle, True)
    
    MsgBox "Completion!", vbInformation
    
    cl.DPUpdate_On
    
End Sub
Sub BtExportMonthlyReportPDF()
    
    cl.DPUpdate_Off
    
    Dim strPath As String, strFile As String
    '파일 내보내기
    strPath = ReplacePath("{tPath}") & "\03_보고파일\02_월간보고\" & [I2] & "\" & [I3] & "\"
    strFile = strPath & [C15] & ".pdf"
    
    Call MkFolder(strPath)
    Call ExportPDF(strFile, , True)
    
    MsgBox "Completion!", vbInformation
    
    cl.DPUpdate_On
End Sub
