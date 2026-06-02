Attribute VB_Name = "tpl_Buttons"
Option Explicit
'<<< 공통 프로시저 시작 ===========================================================================================
Sub BtSearch(Optional ByVal blnEvent As Boolean = True)
    
    On Error GoTo ErrorHandler
    
    Dim dv As DBvar
    With dv
        Dim intReq As Integer
        
        '1. 데이터베이스 ============
        .Alias = [M3]
        
        If .Alias = "" Then
            MsgBox "데이터베이스 예약어를 입력해 주세요!", vbCritical
            Exit Sub
        End If
        
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
        
        .QPlus = [O3]
        
        If .QPlus = "1" Then
            MsgBox "이 페이지의 검색 값은 필수 입니다.", vbCritical
            Exit Sub
        End If
        
        .cstQuery = [P3]
                
        If .cstQuery <> "" Then '커스텀 쿼리 적용
            .Query = .cstQuery
        Else '자동 생성 쿼리 적용
            .Table = [N3]
            
            If .Table = "" Then
                MsgBox "데이터 테이블을 입력하세요!", vbCritical
                Exit Sub
            End If
            
            .Table = IIf(.Type = "엑셀", "[" & .Table & "$]", .Table)
            
            '필드
            Dim boolParens As Boolean
            If .Type = "엑세스" Then boolParens = True
            .arrFields = ReplaceFields(GetFields(, uiColor.TC16), .Token, .Table, boolParens)
            .fields = Join(.arrFields, ", ")
            
            .Query = "SELECT " & .fields & " FROM " & .Table & " " & .QPlus
        End If
        
        .arrQuery = Array(.Query)
        
        '3. 출력범위 ================
        Set .tbl = ActiveSheet.ListObjects(1)
        Set .Target = .tbl.Range(2, 2)
        
        '4. 테이블 초기화 ================
        If blnEvent Then
            cl.DPUpdate_Off
            cl.Event_Off
            cl.Calculate_Off
            cl.sht_UnLock
        End If
        
        Call ClearFiltersInTable(.tbl)
        Call DelTableAllRows(.tbl)
        
        '5. 데이터 로드 ================
        Call SelectQuery(.Target, .arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
        
    End With

    If blnEvent Then
        MsgBox "Completion!", vbInformation
ErrorHandler:
        cl.sht_Lock
        cl.Calculate_On
        cl.Event_On
        cl.DPUpdate_On
    End If
Exit Sub

    If Err.Number <> 0 Then Call HandleError("BtSearch")
End Sub
Sub BtSave()
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    cl.DPUpdate_Off
    cl.Event_Off
    cl.Calculate_Off
    cl.sht_UnLock

    Dim dv As DBvar
    With dv
    
        '<<저장 데이터가 있는 테이블 설정>>
        Dim r, rng As Range
        Dim intColID, _
                intColDel As Integer
        
        On Error Resume Next
        Set .tbl = ActiveSheet.ListObjects(1)
        Set rng = .tbl.ListColumns("선택").DataBodyRange
        
        intColID = intOffset(rng, "ID")
        intColDel = intOffset(rng, "삭제")
        
        If Err.Number <> 0 Then
            MsgBox "저장에 필요한 필드가 누락되어 있습니다!" & vbNewLine & "[ID, REG.OK, 삭제, 선택] 필드가 있는지 확인하세요!", vbCritical
            GoTo ErrorHandler
        End If
        On Error GoTo 0

        Dim intCount As Long
        intCount = Application.WorksheetFunction.CountIf(rng, 1)
        
        If intCount > 0 Then
            Set rng = rng.SpecialCells(xlCellTypeConstants, xlNumbers)
        Else
            MsgBox "선택된 항목이 없습니다!", vbInformation
            GoTo ErrorHandler
        End If

        '1. 데이터베이스 정보 ============
        .Alias = [Q3]
        
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
        
        .Table = [R3]
        
        If .Table = "" Then
            MsgBox "데이터 테이블을 입력하세요!", vbCritical
            GoTo ErrorHandler
        End If
        
        '2. 선택항목 쿼리 변환 ============
        Dim boolParens As Boolean
        boolParens = False
        If .Type = "엑세스" Then boolParens = True
        
        .arrFields = GetFields(uiColor.TC10)
        .arrFields = GetFieldAndType(.arrFields, .Type, .Token, .Table, .File, .Server, .Port, .db, .ID, .PW, boolParens)
        
        If IsArrayEmpty(.arrFields) Then GoTo ErrorHandler
        
        Dim strID As String
        Dim i, j, c As Long
        
        For Each r In rng
            strID = r.Offset(, intColID)
            '쿼리 타입 설정
            Select Case True
            Case r.Offset(, intColDel) = 1 And strID <> "" '삭제
                .QType = "D"
                If .Type = "엑셀" Then
                    Call DelExcelRecQuery(Range(.Table).Worksheet.Name, strID)
                    GoTo pass
                End If
            Case r.Offset(, intColDel) = "" And strID <> "" '수정
                .QType = "U"
            Case r.Offset(, intColDel) = "" And strID = "" '추가
                .QType = "I"
            Case Else
                GoTo passWrong
            End Select
            
            '데이터 수집
            Erase .arrData
            
            For i = 0 To UBound(.arrFields, 2)
                ReDim Preserve .arrData(1, i)
                .arrData(0, i) = .arrFields(0, i)
                .arrData(1, i) = FormatValueForSQLByDBType(r.Offset(, intOffset(r, .arrFields(1, i))), CStr(.arrFields(2, i)))
            Next i
            
            Dim strTable As String
            strTable = IIf(.Type = "엑셀", "[" & .Table & "$]", .Table)
            
            '쿼리 생성
            ReDim Preserve .arrQuery(j)
            .arrQuery(j) = UpsertQuery(.QType, strTable, .arrData, IIf(.QType = "U" Or .QType = "D", "ID = " & strID, ""))
            j = j + 1
pass:
            c = c + 1
passWrong:
                        
            '쿼리 생성 현황 프로그레스
            If Round((c / intCount) * 100) Mod 1 = 0 Then
                Application.StatusBar = "진행 중: " & c & " / " & intCount & _
                                        " (" & Format(c / intCount, "0%") & ")"
                DoEvents ' 간헐적으로만 실행
            End If
        Next r
        
        If IsArrayEmpty(.arrQuery) Then GoTo ErrorHandler
        
        Application.StatusBar = "쿼리생성 완료! 데이터를 전송합니다."
        
        If ExecuteQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW) = False Then GoTo ErrorHandler
        If .Type = "엑셀" Then Call AutoIDs(Range(.Table & "[ID]"))
        
        If ActiveSheet.Name = "필드명매핑" Then
            Call GetFieldMap
            Call GetFieldMapRev
        End If
        
        Call ClearFiltersInTable(.tbl)
        Call DelTableAllRows(.tbl)
        Call BtUpdateDBonThisPage(False)

        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        Application.StatusBar = False
        
        MsgBox c & ".rec Completion!", vbInformation
        
ErrorHandler:
        cl.sht_Lock
        cl.Calculate_On
        cl.Event_On
        cl.DPUpdate_On
        If Err.Number <> 0 Then Call HandleError("BtSave")
    End With

End Sub
'<<< 공통 프로시저 종료 ===========================================================================================

Sub BtSearch_필드명매핑()
    
    Dim dv As DBvar
    With dv
        .Alias = ActiveSheet.cboDB.value
        .Table = ActiveSheet.cboTable.value
        
        If .Alias = "" Or .Table = "" Then
            MsgBox "[데이터베이스], [테이블] 검색 값은 필수 입니다!", vbInformation
            Exit Sub
        End If
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)
                
        Set .tbl = ActiveSheet.ListObjects(1)
        Set .Target = .tbl.ListColumns(4).Range(2, 1)
        
        Dim tblTs As ListObject
        Set tblTs = GetTwbRange("TS_필드명매핑").ListObject
        
        cl.DPUpdate_Off
        cl.Event_Off
        cl.Calculate_Off
        cl.sht_UnLock
        
        Call ClearFiltersInTable(.tbl)
        Call DelTableAllRows(.tbl)
        Call DelTableAllRows(tblTs)
        
        tblTs.ListColumns("예약어").DataBodyRange = .Token
        tblTs.ListColumns("별칭").DataBodyRange = .Alias
        tblTs.ListColumns("데이터베이스").DataBodyRange = .db
        tblTs.ListColumns("파일").DataBodyRange = .File
        tblTs.ListColumns("테이블").DataBodyRange = .Table
        
        If AccessTableExists(.File, .Table) Then
            Dim arr
            Call GetFieldMapRev
            arr = GetFieldNameConnection(.Token, .Table)
            
            If IsArrayEmpty(arr) Then
                MsgBox "[" & .Table & "]테이블의 필드가 존재하지 않습니다!", vbCritical
                GoTo ExitSub
            End If

            .Target.Resize(UBound(arr, 2) + 1, 2) = Application.WorksheetFunction.Transpose(arr)
            
            .tbl.ListColumns("예약어").DataBodyRange = .Token
            .tbl.ListColumns("테이블").DataBodyRange = .Table
            
            Dim strKey As String
            Dim r, rng As Range
            
            Set rng = .tbl.ListColumns("예약어").DataBodyRange
            
            For Each r In rng
                strKey = r & r.Offset(, 1) & r.Offset(, 2)
                If dicPubRev.Exists(strKey) Then r.Offset(, 4).value = dicPubRev(strKey)(1)
            Next

            MsgBox "Completion!", vbInformation
        Else
            MsgBox "[" & .Table & "]테이블이 존재하지 않습니다!" & vbNewLine & "필드를 추가하면 신규로 제작하실 수 있습니다!", vbInformation
        End If

ExitSub:
        cl.sht_Lock
        cl.Calculate_On
        cl.Event_On
        cl.DPUpdate_On
    End With
    
End Sub
Sub BtCheckData()

    Dim dv As DBvar
    With dv
        .Alias = [U3]
        .Table = [v3]
        
        If .Alias = "" Or .Table = "" Then
            MsgBox "[데이터베이스], [테이블] 검색 값은 필수 입니다!", vbInformation
            Exit Sub
        End If
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)
        
        .Query = "SELECT * FROM " & .Table
        .arrQuery = Array(.Query)
        
        Dim celTarget As Range
        Set celTarget = [B9]
        
        cl.DPUpdate_Off
        cl.Event_Off
        cl.sht_UnLock
        
        Call ClearContents(celTarget)
        Call SelectQuery(celTarget, .arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW, , , True)
        celTarget.CurrentRegion.AutoFilter
        
        cl.sht_Lock
        cl.Event_On
        cl.DPUpdate_On
    End With
End Sub

