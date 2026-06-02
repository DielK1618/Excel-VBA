Attribute VB_Name = "am_Table"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Table                                               │
' │  역할 : ListObject(테이블) CRUD, 필터, 정렬, 검색       │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  테이블 조회
' ══════════════════════════════════════════════════════════

' 목적   : 워크북 내 모든 시트의 테이블명 조회
' 인수   : wb - 대상 워크북
' 반환   : Variant - (시트명, 테이블명) 2열 배열
' 예시   : arr = GetAllSheetTableNames(ThisWorkbook)
Public Function GetAllSheetTableNames(ByVal wb As Workbook) As Variant

    Dim ws          As Worksheet
    Dim tbl         As ListObject
    Dim arrResult() As Variant
    Dim i           As Long

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If InStr(tbl.Name, "TS_") = 0 Then
                ReDim Preserve arrResult(1, i)
                arrResult(0, i) = ws.Name
                arrResult(1, i) = tbl.Name
                i = i + 1
            End If
        Next tbl
    Next ws

    GetAllSheetTableNames = Application.Transpose(arrResult)

End Function

' 목적   : 시트 내 테이블명 목록 조회
' 인수   : ws          - 대상 시트
'          strPassName - 제외할 테이블명 패턴 (Like 패턴 사용)
' 반환   : Variant - 테이블명 배열
' 예시   : arr = GetTableNames(Sheet1, "TS_*")
Public Function GetTableNames(ByVal ws As Worksheet, _
                              Optional ByVal strPassName As String = "") As Variant

    Dim tbl         As ListObject
    Dim arrResult() As Variant
    Dim i           As Long

    For Each tbl In ws.ListObjects
        If strPassName <> "" Then
            If tbl.Name Like strPassName Then GoTo NextTbl
        End If
        ReDim Preserve arrResult(i)
        arrResult(i) = tbl.Name
        i = i + 1
NextTbl:
    Next tbl

    GetTableNames = arrResult

End Function

' 목적   : 테이블 각 열 너비 배열 반환
' 인수   : tbl           - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
'          intRound      - 반올림 자릿수
'          sngMultiplier - 배율
' 반환   : Variant - 각 열너비 배열
' 예시   : arr = GetTableColumnsWidth(tbl, 0, 1.2)
Public Function GetTableColumnsWidth(Optional ByVal tbl As ListObject, _
                                     Optional ByVal intRound As Integer = 0, _
                                     Optional ByVal sngMultiplier As Single = 1) As Variant

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim arrResult() As Variant
    Dim r           As Range
    Dim i           As Long

    For Each r In tbl.HeaderRowRange
        ReDim Preserve arrResult(i)
        arrResult(i) = Round(r.Width, intRound) * sngMultiplier
        i = i + 1
    Next r

    GetTableColumnsWidth = arrResult

End Function

' 목적   : 테이블 크기 자동 조정
' 인수   : tbl - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : ResizeTable(tbl)
Public Sub ResizeTable(Optional ByVal tbl As ListObject)

    On Error Resume Next
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    tbl.Resize tbl.Range.CurrentRegion
    On Error GoTo 0

End Sub

' 목적   : 범위가 테이블에 속하는지 확인
' 인수   : rng - 확인할 범위
' 반환   : Boolean - True: 테이블 소속 / False: 아님
' 예시   : If IsTable(Selection) Then ...
Public Function IsTable(ByVal rng As Range) As Boolean
    On Error Resume Next
    IsTable = Not rng.Cells(1, 1).ListObject Is Nothing
    On Error GoTo 0
End Function

' ══════════════════════════════════════════════════════════
'  테이블 행 / 열 조작
' ══════════════════════════════════════════════════════════

' 목적   : 테이블 필터 전체 해제
' 인수   : tbl - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : ClearFiltersInTable(tbl)
Public Sub ClearFiltersInTable(Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    If tbl Is Nothing Then Exit Sub

    On Error Resume Next
    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    On Error GoTo 0

End Sub

' 목적   : 테이블 행 추가
' 인수   : intAddRows   - 추가할 행 수
'          intSelectRow - 삽입 위치 (0: 마지막 행 뒤에)
'          tbl          - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : AddTableRows(3)
Public Sub AddTableRows(ByVal intAddRows As Long, _
                        Optional ByVal intSelectRow As Long = 0, _
                        Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim i As Long
    For i = 1 To intAddRows
        If intSelectRow > 0 Then
            tbl.ListRows.Add (intSelectRow + i)
        Else
            tbl.ListRows.Add AlwaysInsert:=True
        End If
    Next i

End Sub

' 목적   : 테이블 열 추가
' 인수   : intAddColumns   - 추가할 열 수
'          intSelectColumn - 삽입 위치 (0: 마지막 열 뒤에)
'          tbl             - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : AddColumns(2)
Public Sub AddColumns(ByVal intAddColumns As Long, _
                      Optional ByVal intSelectColumn As Long = 0, _
                      Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim i As Long
    For i = 1 To intAddColumns
        If intSelectColumn > 0 Then
            tbl.ListColumns.Add(intSelectColumn + i).Name = "NewCol." & i
        Else
            tbl.ListColumns.Add().Name = "NewCol." & tbl.ListColumns.Count + 1
        End If
    Next i

End Sub

' 목적   : 배열 이름으로 테이블 열 추가
' 인수   : arrColumnNames  - 추가할 열명 배열
'          intSelectColumn - 삽입 위치 (0: 마지막 열 뒤에)
'          tbl             - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : AddArrayColumns(Array("이름", "부서", "직급"), 0, tbl)
Public Sub AddArrayColumns(ByVal arrColumnNames As Variant, _
                           Optional ByVal intSelectColumn As Long = 0, _
                           Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim vntName As Variant
    Dim i       As Long

    For Each vntName In arrColumnNames
        i = i + 1
        If intSelectColumn > 0 Then
            tbl.ListColumns.Add(intSelectColumn + i).Name = vntName
        Else
            tbl.ListColumns.Add().Name = vntName
        End If
    Next vntName

End Sub

' 목적   : 테이블 열 삭제
' 인수   : intSelectColumn - 삭제 시작 열 번호
'          intCount        - 삭제할 열 수 (기본: 1)
'          tbl             - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : DelTableColumns(3, 2, tbl)
Public Sub DelTableColumns(ByVal intSelectColumn As Long, _
                           Optional ByVal intCount As Long = 1, _
                           Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim i As Long
    On Error Resume Next
    For i = intCount To 1 Step -1
        If intSelectColumn + (i - 1) = 1 Then
            tbl.ListColumns(1).Range.Delete
        Else
            tbl.ListColumns(intSelectColumn + (i - 1)).Delete
        End If
    Next i
    On Error GoTo 0

End Sub

' 목적   : 테이블 행 삭제
' 인수   : intSelectRow - 삭제 시작 행 번호
'          intCount     - 삭제할 행 수 (기본: 1)
'          tbl          - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : DelTableRows(2, 3, tbl)
Public Sub DelTableRows(ByVal intSelectRow As Long, _
                        Optional ByVal intCount As Long = 1, _
                        Optional ByVal tbl As ListObject)

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim i As Long
    On Error Resume Next
    For i = intCount To 1 Step -1
        If intSelectRow + (i - 1) = 1 Then
            If tbl.DataBodyRange.Rows.Count > 1 Then
                tbl.ListRows(1).Delete
            Else
                tbl.ListRows(1).Range.SpecialCells(xlCellTypeConstants, 23).ClearContents
            End If
        Else
            tbl.ListRows(intSelectRow + (i - 1)).Delete
        End If
    Next i
    On Error GoTo 0

End Sub

' 목적   : 테이블 전체 행 삭제
' 인수   : tbl - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 예시   : DelTableAllRows(tbl)
Public Sub DelTableAllRows(Optional ByVal tbl As ListObject)

    On Error Resume Next
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim ws      As Worksheet
    Dim rng     As Range
    Dim rngLast As Range

    Set ws = tbl.Range.Worksheet

    If tbl.ListRows.Count > 1 Then
        Set rng = tbl.Range.Offset(2)
        Set rng = rng.Resize(rng.Rows.Count - 2)

        If ws.ListObjects.Count > 1 Then
            rng.Delete Shift:=xlUp
        Else
            rng.EntireRow.Delete
        End If
    End If

    Set rngLast = tbl.ListRows(1).Range
    With rngLast
        If .Cells.Count > 1 Or .HasFormula = True Then
            .SpecialCells(xlCellTypeConstants, 23).ClearContents
        Else
            .ClearContents
        End If
    End With

    On Error GoTo 0

End Sub

' 목적   : 특정 필드에서 조건 일치 행 일괄 삭제
' 인수   : strFieldName     - 검색할 필드명
'          strFilterPattern - Like 조건 문자열
'          tbl              - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
' 반환   : Long - 삭제된 행 수
' 예시   : DelTableFilteredRows("상태", "완료*", tbl)
Public Function DelTableFilteredRows(ByVal strFieldName As String, _
                                     ByVal strFilterPattern As String, _
                                     Optional ByVal tbl As ListObject) As Long

    On Error GoTo ErrHandler

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    If tbl Is Nothing Then Exit Function

    Dim lngFieldCol As Long
    Dim rngDelete   As Range
    Dim lngRow      As Long
    Dim strCellVal  As String

    lngFieldCol = prv_GetFieldColumn(tbl, strFieldName)
    If lngFieldCol = 0 Then
        MsgBox "필드 '" & strFieldName & "'을(를) 찾을 수 없습니다.", _
               vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    For lngRow = tbl.DataBodyRange.Rows.Count To 1 Step -1
        strCellVal = CStr(tbl.DataBodyRange.Cells(lngRow, lngFieldCol).Value)
        If strCellVal Like strFilterPattern Then
            If rngDelete Is Nothing Then
                Set rngDelete = tbl.DataBodyRange.Rows(lngRow)
            Else
                Set rngDelete = Union(rngDelete, tbl.DataBodyRange.Rows(lngRow))
            End If
        End If
    Next lngRow

    If Not rngDelete Is Nothing Then rngDelete.Delete Shift:=xlUp

    Exit Function

ErrHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical, am_Core.AM_NAME

End Function

' ══════════════════════════════════════════════════════════
'  테이블 필터
' ══════════════════════════════════════════════════════════

' 목적   : 테이블 단일 필드 필터 적용
' 인수   : tbl       - 대상 테이블
'          fieldName - 필터 적용 필드명
'          strValue  - 필터 값 ("": 필터 해제)
' 예시   : AutoTableFilter(tbl, "직급", "대리")
Public Sub AutoTableFilter(ByVal tbl As ListObject, _
                           ByVal fieldName As Variant, _
                           ByVal strValue As String)

    Dim intField As Integer
    intField = tbl.ListColumns(fieldName).Index

    If strValue = "" Then
        tbl.Range.AutoFilter intField
    Else
        tbl.Range.AutoFilter intField, strValue
    End If

End Sub

' 목적   : 테이블 필터 다중 값 적용
' 인수   : tbl       - 대상 테이블
'          fieldName - 필터 적용 필드명
'          arrValues - 필터 값 배열
'          blnPart   - True: 부분 일치 (와일드카드 사용)
' 예시   : AutoTableFilter_Arr(tbl, "직급", Array("대리", "과장"))
Public Sub AutoTableFilter_Arr(ByVal tbl As ListObject, _
                               ByVal fieldName As String, _
                               ByVal arrValues As Variant, _
                               Optional ByVal blnPart As Boolean = False)

    On Error Resume Next

    Dim intField       As Integer
    Dim arrWildCards() As String
    Dim i              As Long

    intField = tbl.ListColumns(fieldName).Index

    If blnPart Then
        ReDim arrWildCards(LBound(arrValues) To UBound(arrValues))
        For i = LBound(arrValues) To UBound(arrValues)
            arrWildCards(i) = "*" & arrValues(i) & "*"
        Next i
    Else
        arrWildCards = arrValues
    End If

    tbl.Range.AutoFilter intField, arrWildCards, Operator:=xlFilterValues

    On Error GoTo 0

End Sub

' ══════════════════════════════════════════════════════════
'  테이블 정렬
' ══════════════════════════════════════════════════════════

' 목적   : 테이블 단일 필드 정렬
' 인수   : strTableName  - 테이블명 (Named Range 방식)
'          strFieldName  - 정렬 기준 필드명
'          xlOrder       - 정렬 방향 (기본: 오름차순)
'          xlOrientation - 정렬 방향 (기본: 열 방향)
'          xlHeader      - 헤더 처리 여부 (기본: 있음)
' 예시   : SortTable("T_목록", "이름")
Public Sub SortTable(ByVal strTableName As String, _
                     ByVal strFieldName As String, _
                     Optional ByVal xlOrder As XlSortOrder = xlAscending, _
                     Optional ByVal xlOrientation As XlSortOrientation = xlSortColumns, _
                     Optional ByVal xlHeader As XlYesNoGuess = xlYes)

    Dim tbl As ListObject
    Dim rng As Range
    Set rng = Range(strTableName)
    Set tbl = rng.ListObject

    tbl.Sort.SortFields.Clear
    tbl.Range.Sort tbl.ListColumns(strFieldName).Range, xlOrder, , , , , , xlHeader, , , xlOrientation

End Sub

' 목적   : 사용자 정의 순서로 테이블 정렬
' 인수   : strTableName - 테이블명 (Named Range 방식)
'          strFieldName - 정렬 기준 필드명
'          customList   - 사용자 정의 순서 배열
'          xlOrder      - 정렬 방향 (기본: 오름차순)
' 예시   : SortTableCustomList("T_목록", "지역", Array("서울","부산","대구","울산","광주"))
Public Sub SortTableCustomList(ByVal strTableName As String, _
                               ByVal strFieldName As String, _
                               ByVal customList As Variant, _
                               Optional ByVal xlOrder As XlSortOrder = xlAscending)

    Dim tbl            As ListObject
    Dim rng            As Range
    Dim dicCustomOrder As Object
    Dim colTemp        As ListColumn
    Dim rngSortCol     As Range
    Dim rngTempData    As Range
    Dim arrValues()    As Variant
    Dim arrOrder()     As Variant
    Dim lngRowCount    As Long
    Dim i              As Long
    Dim blnScreen      As Boolean
    Dim xlCalc         As XlCalculation

    blnScreen = Application.ScreenUpdating
    xlCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanUp

    Set dicCustomOrder = CreateObject("Scripting.Dictionary")
    For i = LBound(customList) To UBound(customList)
        dicCustomOrder(customList(i)) = i
    Next i

    Set rng = Range(strTableName)
    Set tbl = rng.ListObject
    Set rngSortCol = tbl.ListColumns(strFieldName).DataBodyRange
    lngRowCount = rngSortCol.Rows.Count

    arrValues = rngSortCol.Value
    ReDim arrOrder(1 To lngRowCount, 1 To 1)

    For i = 1 To lngRowCount
        arrOrder(i, 1) = IIf(dicCustomOrder.Exists(arrValues(i, 1)), _
                             dicCustomOrder(arrValues(i, 1)), 999999)
    Next i

    tbl.ListColumns.Add
    Set colTemp = tbl.ListColumns(tbl.ListColumns.Count)
    Set rngTempData = colTemp.DataBodyRange
    rngTempData.Value = arrOrder

    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=colTemp.Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlOrder, _
                        DataOption:=xlSortNormal
        .Apply
    End With

    colTemp.Delete

CleanUp:
    Application.ScreenUpdating = blnScreen
    Application.Calculation = xlCalc
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "오류 발생: " & Err.Description, vbCritical, am_Core.AM_NAME
    End If

End Sub

' ══════════════════════════════════════════════════════════
'  테이블 값 조작
' ══════════════════════════════════════════════════════════

' 목적   : 테이블 레코드 값 변경
' 인수   : strID        - 검색할 ID 값
'          strFieldName - 변경할 필드명
'          strValue     - 변경할 값
'          tbl          - 대상 테이블 (기본: ActiveSheet 첫번째 테이블)
'          strIDField   - ID 필드명 (기본: "ID")
' 예시   : ChangeTableValue("001", "상태", "완료", tbl)
Public Sub ChangeTableValue(ByVal strID As String, _
                            ByVal strFieldName As String, _
                            ByVal strValue As String, _
                            Optional ByVal tbl As ListObject, _
                            Optional ByVal strIDField As String = "ID")

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim rngFind As Range
    Dim f       As Range

    Set rngFind = tbl.ListColumns(strIDField).DataBodyRange
    Set f = rngFind.Find(strID, , xlValues, xlWhole)

    If Not f Is Nothing Then
        f.Offset(, intOffset(f, strFieldName)).Value = strValue
    End If

End Sub

' 목적   : 테이블 필드 열의 오프셋 계산
' 인수   : CelFrom    - 기준 셀
'          strTgField - 대상 필드명
'          blnTable   - True: 테이블 방식 / False: Named Range 방식
' 반환   : Integer - 열 오프셋 값
' 예시   : intOffset(f, "상태") → 3
Public Function intOffset(ByVal CelFrom As Range, _
                          ByVal strTgField As String, _
                          Optional ByVal blnTable As Boolean = True) As Integer

    On Error Resume Next
    If blnTable Then
        intOffset = Range(CelFrom.ListObject.Name & "[" & strTgField & "]").Column - CelFrom.Column
    Else
        intOffset = Range(strTgField).Column - CelFrom.Column
    End If
    If Err.Number <> 0 Then
        MsgBox "Offset 계산 중 오류입니다!", vbCritical, am_Core.AM_NAME
    End If
    On Error GoTo 0

End Function

' ══════════════════════════════════════════════════════════
'  테이블 검색 (다중 조건)
' ══════════════════════════════════════════════════════════

' 목적   : 다중 조건 검색 후 값 배열 반환
' 인수   : tbl           - 검색할 테이블
'          strPrintField - 반환할 필드명
'          conditions    - 조건 (필드명, 연산자, 값, [AND/OR] 반복)
' 반환   : Variant - 조건에 맞는 값 배열
' 예시   : TblFindVals_MC(tbl, "이름", "나이", ">=", 20)
'          TblFindVals_MC(tbl, "이름", "나이", ">=", 20, "AND", "직급", "=", "대리")
Public Function TblFindVals_MC(ByVal tbl As ListObject, _
                               ByVal strPrintField As String, _
                               ParamArray conditions() As Variant) As Variant

    Dim arrValues()      As Variant
    Dim headerRow        As Range
    Dim i                As Long
    Dim j                As Long
    Dim blnMet           As Boolean
    Dim lngResultCnt     As Long
    Dim lngPrintCol      As Long
    Dim arrCondResults() As Boolean
    Dim lngCondCnt       As Long
    Dim arrLogicOps()    As String
    Dim lngCondIdx       As Long
    Dim arrProcessed()   As Variant

    Set headerRow = tbl.HeaderRowRange

    If UBound(conditions) = 0 And IsArray(conditions(0)) Then
        Dim arrTemp() As Variant
        ReDim arrTemp(LBound(conditions(0)) To UBound(conditions(0)))
        Dim k As Long
        For k = LBound(conditions(0)) To UBound(conditions(0))
            arrTemp(k) = conditions(0)(k)
        Next k
        arrProcessed = arrTemp
    Else
        arrProcessed = conditions
    End If

    If UBound(arrProcessed) < 2 Then
        TblFindVals_MC = Array()
        Exit Function
    End If

    lngPrintCol = prv_GetColumnIndex(headerRow, strPrintField)
    If lngPrintCol = 0 Then
        MsgBox "출력 열 '" & strPrintField & "'을(를) 찾을 수 없습니다.", _
               vbExclamation, am_Core.AM_NAME
        TblFindVals_MC = Array()
        Exit Function
    End If

    lngCondCnt = 0
    ReDim arrCondResults(0 To 100)
    ReDim arrLogicOps(0 To 100)

    j = 0
    Do While j <= UBound(arrProcessed)
        If j + 2 <= UBound(arrProcessed) Then
            lngCondCnt = lngCondCnt + 1
            If j + 3 <= UBound(arrProcessed) Then
                If UCase(Trim(CStr(arrProcessed(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(arrProcessed(j + 3)))) = "OR" Then
                    arrLogicOps(lngCondCnt - 1) = UCase(Trim(CStr(arrProcessed(j + 3))))
                    j = j + 4
                Else
                    arrLogicOps(lngCondCnt - 1) = "AND"
                    j = j + 3
                End If
            Else
                arrLogicOps(lngCondCnt - 1) = "AND"
                j = j + 3
            End If
        Else
            Exit Do
        End If
    Loop

    If lngCondCnt = 0 Then
        TblFindVals_MC = Array()
        Exit Function
    End If

    ReDim Preserve arrCondResults(0 To lngCondCnt - 1)
    ReDim Preserve arrLogicOps(0 To lngCondCnt - 1)

    For i = 1 To tbl.ListRows.Count

        j = 0
        lngCondIdx = 0

        Do While lngCondIdx < lngCondCnt
            Dim strColName  As String
            Dim strOperator As String
            Dim vntValue    As Variant
            Dim lngColIdx   As Long
            Dim vntCell     As Variant

            strColName = CStr(arrProcessed(j))
            strOperator = CStr(arrProcessed(j + 1))
            vntValue = arrProcessed(j + 2)

            lngColIdx = prv_GetColumnIndex(headerRow, strColName)
            If lngColIdx = 0 Then
                arrCondResults(lngCondIdx) = False
            Else
                vntCell = tbl.ListRows(i).Range.Cells(1, lngColIdx).Value
                arrCondResults(lngCondIdx) = prv_EvaluateCondition(vntCell, strOperator, vntValue)
            End If

            If j + 3 <= UBound(arrProcessed) Then
                If UCase(Trim(CStr(arrProcessed(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(arrProcessed(j + 3)))) = "OR" Then
                    j = j + 4
                Else
                    j = j + 3
                End If
            Else
                j = j + 3
            End If

            lngCondIdx = lngCondIdx + 1
        Loop

        blnMet = arrCondResults(0)
        For lngCondIdx = 1 To lngCondCnt - 1
            If arrLogicOps(lngCondIdx - 1) = "AND" Then
                blnMet = blnMet And arrCondResults(lngCondIdx)
            ElseIf arrLogicOps(lngCondIdx - 1) = "OR" Then
                blnMet = blnMet Or arrCondResults(lngCondIdx)
            End If
        Next lngCondIdx

        If blnMet Then
            ReDim Preserve arrValues(lngResultCnt)
            arrValues(lngResultCnt) = tbl.ListRows(i).Range.Cells(1, lngPrintCol).Value
            lngResultCnt = lngResultCnt + 1
        End If

    Next i

    TblFindVals_MC = IIf(lngResultCnt > 0, arrValues, Array())

End Function

' 목적   : 다중 조건 검색 후 단일 값 반환
' 인수   : tbl           - 검색할 테이블
'          strPrintField - 반환할 필드명
'          lngIndex      - 반환할 결과 순번 (1부터 시작)
'          conditions    - 조건 (TblFindVals_MC 와 동일)
' 반환   : Variant - 조건에 맞는 lngIndex 번째 값
' 예시   : TblFindVal_One(tbl, "이름", 1, "나이", ">=", 20)
Public Function TblFindVal_One(ByVal tbl As ListObject, _
                               ByVal strPrintField As String, _
                               ByVal lngIndex As Long, _
                               ParamArray conditions() As Variant) As Variant

    Dim arrCond()  As Variant
    Dim arrResult  As Variant
    Dim i          As Long

    ReDim arrCond(LBound(conditions) To UBound(conditions))
    For i = LBound(conditions) To UBound(conditions)
        arrCond(i) = conditions(i)
    Next i

    arrResult = TblFindVals_MC(tbl, strPrintField, arrCond)

    If IsArray(arrResult) And UBound(arrResult) >= lngIndex - 1 Then
        TblFindVal_One = arrResult(lngIndex - 1)
    Else
        TblFindVal_One = ""
    End If

End Function

' 목적   : 다중 조건 검색 후 Range 반환
' 인수   : targetTable   - 검색할 테이블
'          outputColumns - 반환할 열명 (문자열 또는 배열)
'          conditions    - 조건 (TblFindVals_MC 와 동일)
' 반환   : Range - 조건에 맞는 행 범위
' 예시   : TblFindRng_MC(tbl, "이름", "나이", ">=", 20)
Public Function TblFindRng_MC(ByVal targetTable As ListObject, _
                              ByVal outputColumns As Variant, _
                              ParamArray conditions() As Variant) As Range

    Dim rngResult        As Range
    Dim rngOutput        As Range
    Dim headerRow        As Range
    Dim arrOutColIdx()   As Long
    Dim lngOutColCnt     As Long
    Dim arrCondResults() As Boolean
    Dim lngCondCnt       As Long
    Dim arrLogicOps()    As String
    Dim arrProcessed()   As Variant
    Dim i                As Long
    Dim j                As Long
    Dim lngCondIdx       As Long
    Dim blnMet           As Boolean

    Set headerRow = targetTable.HeaderRowRange

    If VarType(outputColumns) = vbString Then
        ReDim arrOutColIdx(0 To 0)
        arrOutColIdx(0) = prv_GetColumnIndex(headerRow, CStr(outputColumns))
        If arrOutColIdx(0) = 0 Then
            MsgBox "출력 열 '" & outputColumns & "'을(를) 찾을 수 없습니다.", _
                   vbExclamation, am_Core.AM_NAME
            Set TblFindRng_MC = Nothing
            Exit Function
        End If
        lngOutColCnt = 1
    Else
        ReDim arrOutColIdx(0 To UBound(outputColumns))
        For i = 0 To UBound(outputColumns)
            arrOutColIdx(i) = prv_GetColumnIndex(headerRow, CStr(outputColumns(i)))
            If arrOutColIdx(i) = 0 Then
                MsgBox "출력 열 '" & outputColumns(i) & "'을(를) 찾을 수 없습니다.", _
                       vbExclamation, am_Core.AM_NAME
                Set TblFindRng_MC = Nothing
                Exit Function
            End If
        Next i
        lngOutColCnt = UBound(outputColumns) + 1
    End If

    If UBound(conditions) = 0 And IsArray(conditions(0)) Then
        Dim arrTemp() As Variant
        ReDim arrTemp(LBound(conditions(0)) To UBound(conditions(0)))
        Dim k As Long
        For k = LBound(conditions(0)) To UBound(conditions(0))
            arrTemp(k) = conditions(0)(k)
        Next k
        arrProcessed = arrTemp
    Else
        arrProcessed = conditions
    End If

    If UBound(arrProcessed) < 2 Then
        Set TblFindRng_MC = Nothing
        Exit Function
    End If

    lngCondCnt = 0
    ReDim arrCondResults(0 To 100)
    ReDim arrLogicOps(0 To 100)

    j = 0
    Do While j <= UBound(arrProcessed)
        If j + 2 <= UBound(arrProcessed) Then
            lngCondCnt = lngCondCnt + 1
            If j + 3 <= UBound(arrProcessed) Then
                If UCase(Trim(CStr(arrProcessed(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(arrProcessed(j + 3)))) = "OR" Then
                    arrLogicOps(lngCondCnt - 1) = UCase(Trim(CStr(arrProcessed(j + 3))))
                    j = j + 4
                Else
                    arrLogicOps(lngCondCnt - 1) = "AND"
                    j = j + 3
                End If
            Else
                arrLogicOps(lngCondCnt - 1) = "AND"
                j = j + 3
            End If
        Else
            Exit Do
        End If
    Loop

    If lngCondCnt = 0 Then
        Set TblFindRng_MC = Nothing
        Exit Function
    End If

    ReDim Preserve arrCondResults(0 To lngCondCnt - 1)
    ReDim Preserve arrLogicOps(0 To lngCondCnt - 1)

    For i = 1 To targetTable.ListRows.Count

        j = 0
        lngCondIdx = 0

        Do While lngCondIdx < lngCondCnt
            Dim strColName2  As String
            Dim strOperator2 As String
            Dim vntValue2    As Variant
            Dim lngColIdx2   As Long
            Dim vntCell2     As Variant

            strColName2 = CStr(arrProcessed(j))
            strOperator2 = CStr(arrProcessed(j + 1))
            vntValue2 = arrProcessed(j + 2)

            lngColIdx2 = prv_GetColumnIndex(headerRow, strColName2)
            If lngColIdx2 = 0 Then
                arrCondResults(lngCondIdx) = False
            Else
                vntCell2 = targetTable.ListRows(i).Range.Cells(1, lngColIdx2).Value
                arrCondResults(lngCondIdx) = prv_EvaluateCondition(vntCell2, strOperator2, vntValue2)
            End If

            If j + 3 <= UBound(arrProcessed) Then
                If UCase(Trim(CStr(arrProcessed(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(arrProcessed(j + 3)))) = "OR" Then
                    j = j + 4
                Else
                    j = j + 3
                End If
            Else
                j = j + 3
            End If

            lngCondIdx = lngCondIdx + 1
        Loop

        blnMet = arrCondResults(0)
        For lngCondIdx = 1 To lngCondCnt - 1
            If arrLogicOps(lngCondIdx - 1) = "AND" Then
                blnMet = blnMet And arrCondResults(lngCondIdx)
            ElseIf arrLogicOps(lngCondIdx - 1) = "OR" Then
                blnMet = blnMet Or arrCondResults(lngCondIdx)
            End If
        Next lngCondIdx

        If blnMet Then
            For j = 0 To lngOutColCnt - 1
                Set rngOutput = targetTable.ListRows(i).Range.Cells(1, arrOutColIdx(j))
                If rngResult Is Nothing Then
                    Set rngResult = rngOutput
                Else
                    Set rngResult = Union(rngResult, rngOutput)
                End If
            Next j
        End If

    Next i

    Set TblFindRng_MC = rngResult

End Function

' ══════════════════════════════════════════════════════════
'  내부 전용 함수 (Private)
' ══════════════════════════════════════════════════════════

' 목적   : 필드명으로 열 번호 반환 (내부 전용)
Private Function prv_GetFieldColumn(ByVal tbl As ListObject, _
                                    ByVal strFieldName As String) As Long

    Dim lngCol As Long
    For lngCol = 1 To tbl.ListColumns.Count
        If StrComp(tbl.ListColumns(lngCol).Name, strFieldName, vbTextCompare) = 0 Then
            prv_GetFieldColumn = lngCol
            Exit Function
        End If
    Next lngCol
    prv_GetFieldColumn = 0

End Function

' 목적   : 헤더행으로 열 인덱스 반환 (내부 전용)
Private Function prv_GetColumnIndex(ByVal headerRow As Range, _
                                    ByVal strColName As String) As Long

    Dim cel As Range
    For Each cel In headerRow.Cells
        If LCase(Trim(cel.Value)) = LCase(Trim(strColName)) Then
            prv_GetColumnIndex = cel.Column - headerRow.Column + 1
            Exit Function
        End If
    Next cel
    prv_GetColumnIndex = 0

End Function

' 목적   : 조건 평가 (내부 전용)
Private Function prv_EvaluateCondition(ByVal vntCell As Variant, _
                                       ByVal strOperator As String, _
                                       ByVal vntValue As Variant) As Boolean

    On Error Resume Next

    Select Case LCase(Trim(strOperator))
        Case "=":    prv_EvaluateCondition = (vntCell = vntValue)
        Case "<>":   prv_EvaluateCondition = (vntCell <> vntValue)
        Case ">":    prv_EvaluateCondition = (vntCell > vntValue)
        Case ">=":   prv_EvaluateCondition = (vntCell >= vntValue)
        Case "<":    prv_EvaluateCondition = (vntCell < vntValue)
        Case "<=":   prv_EvaluateCondition = (vntCell <= vntValue)
        Case "like": prv_EvaluateCondition = (vntCell Like vntValue)
        Case Else:   prv_EvaluateCondition = False
    End Select

    If Err.Number <> 0 Then
        prv_EvaluateCondition = False
        Err.Clear
    End If

    On Error GoTo 0

End Function
