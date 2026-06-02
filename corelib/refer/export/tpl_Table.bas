Attribute VB_Name = "tpl_Table"
Option Explicit
Function GetAllSheetTableNames() As Variant

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim arrResult()
    Dim i As Long
        
    For Each ws In ThisWorkbook.Worksheets
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
Function GetTableNames(Optional ByVal sht As Worksheet, _
                                                Optional strPassName As String) As Variant

    Dim tbl As ListObject
    Dim arrResult()
    Dim i As Long
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    For Each tbl In sht.ListObjects
    
        If tbl.Name Like strPassName Then GoTo pass
        
            ReDim Preserve arrResult(i)
            arrResult(i) = tbl.Name
            i = i + 1
pass:
    Next tbl
    
    GetTableNames = arrResult
    
End Function
Sub ClearFiltersInTable(Optional ByVal tbl As ListObject) '<<< 테이블 필터 제거 >>>
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1) ' 시트 이름을 수정하여 원하는 시트로 변경
    
    If Not tbl Is Nothing Then
        On Error Resume Next
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
        On Error GoTo 0
    End If

End Sub
Sub AddTableRows(ByVal intAddRows As Long, _
                                    Optional ByVal intSelectRow As Long = 0, _
                                    Optional ByVal tbl As ListObject) '<<< 테이블 행 추가, intSelectRow 값을 입력하지 않으면 마지막행 다음으로 삽입 >>>
    
    Dim i As Long
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    For i = 1 To intAddRows
        If intSelectRow > 0 Then
            tbl.ListRows.Add (intSelectRow + i)
        Else
            tbl.ListRows.Add AlwaysInsert:=True
        End If
    Next
    
End Sub
Sub AddColumns(ByVal intAddColumns As Long, _
                                Optional ByVal intSelectColumn As Long = 0, _
                                Optional ByVal tbl As ListObject) '<<< 테이블 열 추가, intSelectColumn 값을 입력하지 않으면 마지막 열 다음으로 삽입 >>>

    Dim i As Long

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    For i = 1 To intAddColumns
        If intSelectColumn > 0 Then
            tbl.ListColumns.Add(intSelectColumn + i).Name = "NewCol." & i ' 새 열 이름을 여기에 지정하세요
        Else
            tbl.ListColumns.Add().Name = "NewCol." & tbl.ListColumns.count + 1 ' 새 열 이름을 여기에 지정하세요
        End If
    Next

End Sub
Sub AddArrayColumns(ByVal arrColumnNames As Variant, _
                                            Optional ByVal intSelectColumn As Long = 0, _
                                            Optional ByVal tbl As ListObject) '<<< 테이블 열 추가, intSelectColumn 값을 입력하지 않으면 마지막 열 다음으로 삽입 >>>
    
    Dim varName As Variant
    Dim i As Long
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    For Each varName In arrColumnNames
        i = i + 1
        If intSelectColumn > 0 Then
            tbl.ListColumns.Add(intSelectColumn + i).Name = varName ' 새 열 이름을 여기에 지정하세요
        Else
            tbl.ListColumns.Add().Name = varName ' 새 열 이름을 여기에 지정하세요
        End If
    Next
    
End Sub
Sub DelTableColumns(ByVal intSelectColumn As Long, _
                    Optional ByVal intCount As Long = 1, _
                    Optional ByVal tbl As ListObject) '<<< 테이블 열 삭제 >>>
    
    Dim i As Long
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    On Error Resume Next
    For i = intCount To 1 Step -1
        If intSelectColumn + (i - 1) = 1 Then
            tbl.ListColumns(1).Range.Delete
        Else
            tbl.ListColumns(intSelectColumn + (i - 1)).Delete
        End If
    Next
    On Error GoTo 0
    
End Sub
Sub DelTableRows(ByVal intSelectRow As Long, _
                                    Optional ByVal intCount As Long = 1, _
                                    Optional ByVal tbl As ListObject) '<<< 테이블 행 추가 >>>
    
    Dim i As Long
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    On Error Resume Next
    For i = intCount To 1 Step -1
        If intSelectRow + (i - 1) = 1 Then
            If tbl.DataBodyRange.Rows.count > 1 Then
                tbl.ListRows(intSelectRow + (i - 1)).Delete
            Else
                tbl.ListRows(1).Range.SpecialCells(xlCellTypeConstants, 23).ClearContents
            End If
        Else
            tbl.ListRows(intSelectRow + (i - 1)).Delete
        End If
    Next
    On Error GoTo 0
    
End Sub
Sub DelTableAllRows(Optional ByVal tbl As ListObject)  '<<< 테이블 행 전체 삭제 >>>

    On Error Resume Next
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim ws As Worksheet
    Set ws = tbl.Range.Worksheet


    If tbl.ListRows.count > 1 Then
        Dim rng As Range

        Set rng = tbl.Range.Offset(2)
        Set rng = rng.Resize(rng.Rows.count - 2)

        If ws.ListObjects.count > 1 Then
            rng.Delete Shift:=xlUp
        Else
            rng.EntireRow.Delete
        End If
    End If
    
    Dim rngLastRow As Range
    Set rngLastRow = tbl.ListRows(1).Range
    With rngLastRow
    If .Cells.count > 1 Or .HasFormula = True Then
        tbl.ListRows(1).Range.SpecialCells(xlCellTypeConstants, 23).ClearContents
    Else
        tbl.ListRows(1).Range.ClearContents
    End If
    End With
End Sub
Public Function DelTableFilteredRows(ByVal strFieldName As String, _
                                                                     ByVal strFilterPattern As String, _
                                                                    Optional ByVal tbl As ListObject) As Long '특정 필드에서 패턴이 일치하는 값이 있는 행을 찾아 일괄 삭제
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim fieldColumn As Long
    Dim deleteRanges As Range
    Dim dataRow As Long
    Dim cellValue As String
    
    'Get table by name
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    If tbl Is Nothing Then
        MsgBox "테이블 '" & tbl.Name & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Function
    End If
    
    '필드 열 번호 찾기
    fieldColumn = GetFieldColumn(tbl, strFieldName)
    If fieldColumn = 0 Then
        MsgBox "필드 '" & strFieldName & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Function
    End If
    
    '테이블의 마지막 행부터 첫 데이터 행까지 역순으로 확인
    For dataRow = tbl.DataBodyRange.Rows.count To 1 Step -1
        '셀 값을 문자열로 변환
        cellValue = CStr(tbl.DataBodyRange.Cells(dataRow, fieldColumn).value)
        
        'LIKE 연산자를 사용하여 패턴 매칭
        If cellValue Like strFilterPattern Then
            If deleteRanges Is Nothing Then
                Set deleteRanges = tbl.DataBodyRange.Rows(dataRow)
            Else
                Set deleteRanges = Union(deleteRanges, tbl.DataBodyRange.Rows(dataRow))
            End If
        End If
    Next dataRow
    
    '한 번에 모든 행 삭제
    If Not deleteRanges Is Nothing Then
        deleteRanges.Delete Shift:=xlUp
    End If
    
CleanUp:
    Exit Function
    
ErrorHandler:
    MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical
    Resume CleanUp
End Function
Function GetFieldColumn(ByVal tbl As ListObject, ByVal strFieldName As String) As Long
    Dim col As Long
    
    For col = 1 To tbl.ListColumns.count
        If StrComp(tbl.ListColumns(col).Name, strFieldName, vbTextCompare) = 0 Then
            GetFieldColumn = col
            Exit Function
        End If
    Next col
    
    GetFieldColumn = 0
End Function
Sub AutoTableFilter(ByVal tbl As ListObject, ByVal fieldName, strValue As String)  '<<< 단일 필터 >>>

    Dim intField As Integer
    
    intField = tbl.ListColumns(fieldName).Index
    
    If strValue = "" Then
        tbl.Range.AutoFilter intField
    Else
        tbl.Range.AutoFilter intField, strValue
    End If
    
End Sub
Sub AutoTableFilter_Arr(ByVal tbl As ListObject, _
                                            ByVal fieldName As String, _
                                            ByVal arrValues As Variant, _
                                            Optional ByVal blnPart As Boolean = False) '<<< 필터 다중 값 >>>
    On Error Resume Next
    Dim intField As Integer
    Dim i As Long
    Dim arrWildCards() As String
    
    ' 필드 인덱스 가져오기
    intField = tbl.ListColumns(fieldName).Index
    
    If blnPart Then
        ' arrValues를 와일드카드 적용 배열로 변환
        ReDim arrWildCards(LBound(arrValues) To UBound(arrValues))
        For i = LBound(arrValues) To UBound(arrValues)
            arrWildCards(i) = "*" & arrValues(i) & "*"
        Next i
    Else
        arrWildCards = arrValues
    End If
    
    ' 필터 적용 (부분 일치)
    tbl.Range.AutoFilter intField, arrWildCards, operator:=xlFilterValues
    
End Sub
Sub ChangeTableValue(ByVal strID As String, _
                                            ByVal strFieldName As String, _
                                            ByVal strValue As String, _
                                            Optional ByVal tbl As ListObject, _
                                            Optional ByVal strIDFieldName As String = "ID") '<<<테이블 레코드 값 수정, Change 이벤트와 함께 쓸 수 있음 >>>

    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim f, rngFind As Range
    Set rngFind = tbl.ListColumns(strIDFieldName).DataBodyRange
    Set f = rngFind.Find(strID, , xlValues, xlWhole)
    
    If Not f Is Nothing Then
            f.Offset(, intOffset(f, strFieldName)).value = strValue
    End If
    
End Sub
Function IsTable(ByVal rng As Range) As Boolean
    '>>> 범위의 첫 셀이 테이블에 속하는지 확인 >>>
    On Error Resume Next
    IsTable = Not rng.Cells(1, 1).ListObject Is Nothing
    On Error GoTo 0
End Function
Function intOffset(ByVal CelFrom As Range, _
                                  ByVal strTgField As String, _
                                  Optional blnTable As Boolean = True) As Integer  '<<<테이블 OFFSET 인덱스 번호 >>>
    
    On Error Resume Next
    If blnTable Then
        intOffset = Range(CelFrom.ListObject.Name & "[" & strTgField & "]").Column - CelFrom.Column
    Else
        intOffset = Range(strTgField).Column - CelFrom.Column
    End If
    If Err.Number <> 0 Then
        MsgBox "Offset 설정 값 오류입니다!", vbCritical
        Exit Function
    End If
    On Error GoTo 0

End Function
Sub ValidationList(ByVal rngTarget As Range, _
                                    ByVal arrValues As Variant) '<<< 목록상자 >>>
        
        Dim strValues As String

        strValues = Join(arrValues, ",")
        
        With rngTarget.Validation
            .Delete

        If strValues <> "" Then .Add xlValidateList, xlValidAlertStop, xlBetween, strValues
        End With
        
End Sub
Function FindRange(ByVal rngFind As Range, _
                                        ByVal strValue As String, _
                                        Optional ByVal blnAddRow As Boolean = True, _
                                        Optional ByVal LookAt As XlLookAt = xlWhole) As Range '<<< Find >>>
                                        
    On Error Resume Next
    Dim f As Range
    
    Set f = rngFind.Find(strValue, , xlValues, LookAt)
    
    If Not f Is Nothing Then
        Set FindRange = f
    Else
        If blnAddRow Then
            Set f = rngFind.Worksheet.Cells(Rows.count, rngFind.Column).End(xlUp)
            Set f = IIf(f = "", f, f.Offset(1))
            
            Set FindRange = f
        Else
        
            Set FindRange = Nothing
        
        End If
    End If
    
End Function
Sub SortTable(ByVal strTableName As String, _
                            ByVal strFieldName As String, _
                            Optional ByVal xlOrder As XlSortOrder = xlAscending, _
                            Optional ByVal xlOrientation As XlSortOrientation = xlSortColumns, _
                            Optional ByVal xlHearder As XlYesNoGuess = xlYes) '<<< 정렬 >>
    
    Dim tbl As ListObject
    Dim rng As Range
    Set rng = Range(strTableName)
    Set tbl = rng.ListObject
    
    tbl.Sort.SortFields.Clear
    tbl.Range.Sort tbl.ListColumns(strFieldName).Range, xlOrder, , , , , , xlHearder, , , xlOrientation

End Sub
Sub SortTableCustomList(ByVal strTableName As String, _
                        ByVal strFieldName As String, _
                        ByVal customList As Variant, _
                        Optional ByVal xlOrder As XlSortOrder = xlAscending)

    Dim tbl         As ListObject
    Dim rng         As Range
    Dim i           As Long
    Dim dicCustomOrder  As Object

    '── 성능 최적화 설정 ──────────────────────────────
    Dim blnScreenUpdate As Boolean
    Dim xlCalcMode      As XlCalculation

    blnScreenUpdate = Application.ScreenUpdating
    xlCalcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanUp   '오류 발생 시 원상복구 보장

    '── Dictionary 생성 및 사용자 정의 목록 저장 ──────
    Set dicCustomOrder = CreateObject("Scripting.Dictionary")
    For i = LBound(customList) To UBound(customList)
        dicCustomOrder(customList(i)) = i   ' .Add 대신 직접 할당 (중복 오류 방지)
    Next i

    '── 테이블 및 정렬 열 설정 ───────────────────────
    Set rng = Range(strTableName)
    Set tbl = rng.ListObject

    Dim rngSortCol  As Range
    Dim lngRowCount As Long

    Set rngSortCol = tbl.ListColumns(strFieldName).DataBodyRange  ' 헤더 제외 데이터 영역
    lngRowCount = rngSortCol.Rows.count

    '── ★ 핵심: 배열로 일괄 읽기 ───────────────────
    Dim arrValues() As Variant
    Dim arrOrder()  As Variant

    arrValues = rngSortCol.value                ' 셀 → 배열 (1회 I/O)
    ReDim arrOrder(1 To lngRowCount, 1 To 1)    ' 쓰기용 배열 초기화

    For i = 1 To lngRowCount
        If dicCustomOrder.Exists(arrValues(i, 1)) Then
            arrOrder(i, 1) = dicCustomOrder(arrValues(i, 1))
        Else
            arrOrder(i, 1) = 999999             ' 목록 외 항목 → 맨 뒤
        End If
    Next i

    '── 임시 열 추가 및 배열로 일괄 쓰기 ────────────
    tbl.ListColumns.Add

    Dim colTemp     As ListColumn
    Dim rngTempData As Range

    Set colTemp = tbl.ListColumns(tbl.ListColumns.count)
    Set rngTempData = colTemp.DataBodyRange

    rngTempData.value = arrOrder                ' 배열 → 셀 (1회 I/O)

    '── 정렬 실행 ────────────────────────────────────
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=colTemp.Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlOrder, _
                        DataOption:=xlSortNormal
        .Apply
    End With

    '── 임시 열 삭제 ─────────────────────────────────
    colTemp.Delete

CleanUp:
    '── 성능 설정 원상복구 ────────────────────────────
    Application.ScreenUpdating = blnScreenUpdate
    Application.Calculation = xlCalcMode
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "오류 발생: " & Err.Description, vbCritical
    End If

End Sub
Function GetTableColumnsWidth(Optional ByVal tbl As ListObject, _
                                                            Optional intRound As Integer = 0, _
                                                            Optional sngMultiplier As Single = 1) As Variant
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    Dim arrResult()
    Dim r As Range
    Dim i As Long
    
    For Each r In tbl.HeaderRowRange
        ReDim Preserve arrResult(i)
        arrResult(i) = Round(r.Width, intRound) * sngMultiplier
        
        i = i + 1
    Next
    
    GetTableColumnsWidth = arrResult
    
End Function
Sub ResizeTable(Optional ByVal tbl As ListObject)
    On Error Resume Next
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)

    Dim strtblAddress As String
    
    strtblAddress = tbl.Range.CurrentRegion.Address
    tbl.Resize Range(strtblAddress)
    On Error GoTo 0
End Sub


'============================================================================
' 테이블 검색 함수 모음 (OR/AND 연산자 지원)
'============================================================================

Public Function TblFindVals_MC(ByVal tbl As ListObject, _
                               ByVal strPrintField As String, _
                               ParamArray conditions() As Variant) As Variant
'----------------------------------------------------------------------------
' 목적: 테이블에서 여러 조건(AND/OR)에 맞는 값들을 찾아 반환
' 매개변수:
'   - tbl: 검색할 테이블
'   - strPrintField: 출력할 열 이름
'   - conditions: 검색 조건들 (열이름, 연산자, 값, [논리연산자] 순서)
'              또는 배열 형태로 전달 가능
' 반환: 조건에 맞는 값들의 배열
' 예시:
'   TblFindVals_MC(tbl, "이름", "나이", ">=", 20)
'   TblFindVals_MC(tbl, "이름", "나이", ">=", 20, "AND", "부서", "=", "개발")
'   TblFindVals_MC(tbl, "이름", "부서", "=", "개발", "OR", "부서", "=", "영업")
'   TblFindVals_MC(tbl, "이름", Array("부서", "=", "개발", "OR", "부서", "=", "영업"))
'----------------------------------------------------------------------------
    
    Dim arrValues()           ' 결과를 저장할 배열
    Dim headerRow As Range    ' 테이블 헤더 행
    Dim i As Long, j As Long  ' 반복문 인덱스
    Dim conditionMet As Boolean     ' 조건 충족 여부
    Dim resultCount As Long   ' 결과 개수
    Dim printColIndex As Long ' 출력 열 인덱스
    Dim conditionResults() As Boolean
    Dim conditionCount As Long
    Dim logicOps() As String
    Dim condIdx As Long
    Dim processedConditions() As Variant
    
    ' 헤더 행 설정
    Set headerRow = tbl.HeaderRowRange
    
    ' 배열로 전달된 경우 처리
    If UBound(conditions) = 0 And IsArray(conditions(0)) Then
        Dim tempArr() As Variant
        ReDim tempArr(LBound(conditions(0)) To UBound(conditions(0)))
        Dim k As Long
        For k = LBound(conditions(0)) To UBound(conditions(0))
            tempArr(k) = conditions(0)(k)
        Next k
        processedConditions = tempArr
    Else
        processedConditions = conditions
    End If
    
    ' 조건이 없는 경우 처리
    If UBound(processedConditions) < 2 Then
        TblFindVals_MC = Array()
        Exit Function
    End If
    
    ' 출력 열 인덱스 찾기
    printColIndex = GetColumnIndex(headerRow, strPrintField)
    If printColIndex = 0 Then
        MsgBox "출력 열 '" & strPrintField & "'을(를) 찾을 수 없습니다.", vbExclamation
        TblFindVals_MC = Array()
        Exit Function
    End If
    
    ' 조건 파싱: 조건 개수와 논리 연산자 추출
    conditionCount = 0
    ReDim conditionResults(0 To 100)
    ReDim logicOps(0 To 100)
    
    j = 0
    Do While j <= UBound(processedConditions)
        If j + 2 <= UBound(processedConditions) Then
            conditionCount = conditionCount + 1
            
            ' 다음 요소가 논리 연산자인지 확인
            If j + 3 <= UBound(processedConditions) Then
                If UCase(Trim(CStr(processedConditions(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(processedConditions(j + 3)))) = "OR" Then
                    logicOps(conditionCount - 1) = UCase(Trim(CStr(processedConditions(j + 3))))
                    j = j + 4
                Else
                    logicOps(conditionCount - 1) = "AND"
                    j = j + 3
                End If
            Else
                logicOps(conditionCount - 1) = "AND"
                j = j + 3
            End If
        Else
            Exit Do
        End If
    Loop
    
    If conditionCount > 0 Then
        ReDim Preserve conditionResults(0 To conditionCount - 1)
        ReDim Preserve logicOps(0 To conditionCount - 1)
    Else
        TblFindVals_MC = Array()
        Exit Function
    End If
    
    resultCount = 0
    
    ' 각 행별로 검색
    For i = 1 To tbl.ListRows.count
        
        ' 각 조건 평가
        j = 0
        condIdx = 0
        
        Do While condIdx < conditionCount
            Dim colName As String
            Dim operator As String
            Dim value As Variant
            Dim colIndex As Long
            Dim cellValue As Variant
            
            ' 조건 파라미터 추출
            colName = CStr(processedConditions(j))
            operator = CStr(processedConditions(j + 1))
            value = processedConditions(j + 2)
            
            ' 열 인덱스 찾기
            colIndex = GetColumnIndex(headerRow, colName)
            If colIndex = 0 Then
                conditionResults(condIdx) = False
            Else
                ' 셀 값 가져오기
                cellValue = tbl.ListRows(i).Range.Cells(1, colIndex).value
                
                ' 조건 비교
                conditionResults(condIdx) = EvaluateCondition(cellValue, operator, value)
            End If
            
            ' 다음 조건으로 이동
            If j + 3 <= UBound(processedConditions) Then
                If UCase(Trim(CStr(processedConditions(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(processedConditions(j + 3)))) = "OR" Then
                    j = j + 4
                Else
                    j = j + 3
                End If
            Else
                j = j + 3
            End If
            
            condIdx = condIdx + 1
        Loop
        
        ' 논리 연산자를 사용하여 최종 결과 계산
        conditionMet = conditionResults(0)
        For condIdx = 1 To conditionCount - 1
            If logicOps(condIdx - 1) = "AND" Then
                conditionMet = conditionMet And conditionResults(condIdx)
            ElseIf logicOps(condIdx - 1) = "OR" Then
                conditionMet = conditionMet Or conditionResults(condIdx)
            End If
        Next condIdx
        
        ' 조건에 맞는 셀 값을 결과 배열에 추가
        If conditionMet Then
            ReDim Preserve arrValues(resultCount)
            arrValues(resultCount) = tbl.ListRows(i).Range.Cells(1, printColIndex).value
            resultCount = resultCount + 1
        End If
    Next i
    
    ' 결과 반환
    If resultCount > 0 Then
        TblFindVals_MC = arrValues
    Else
        TblFindVals_MC = Array()
    End If
End Function
Public Function TblFindVal_One(ByVal tbl As ListObject, _
                               ByVal strPrintField As String, _
                               ByVal lngIndex As Long, _
                               ParamArray conditions() As Variant) As Variant
    Dim arrResult As Variant
    
    ' 기존 함수 호출 시 conditions 재구성 필요
    Dim arrCond() As Variant
    Dim i As Long
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

Public Function TblFindRng_MC(ByVal targetTable As ListObject, _
                              ByVal outputColumns As Variant, _
                              ParamArray conditions() As Variant) As Range
'----------------------------------------------------------------------------
' 목적: 테이블에서 여러 조건(AND/OR)에 맞는 셀들의 범위를 반환
' 매개변수:
'   - targetTable: 검색할 테이블
'   - outputColumns: 출력할 열 이름(단일 문자열 또는 문자열 배열)
'   - conditions: 검색 조건들 (열이름, 연산자, 값, [논리연산자] 순서)
'              또는 배열 형태로 전달 가능
' 반환: 조건에 맞는 셀들의 Range 객체
' 예시:
'   TblFindRng_MC(tbl, "이름", "나이", ">=", 20)
'   TblFindRng_MC(tbl, "이름", "나이", ">=", 20, "AND", "부서", "=", "개발")
'   TblFindRng_MC(tbl, Array("이름", "부서"), "나이", ">=", 20, "OR", "부서", "=", "개발")
'   TblFindRng_MC(tbl, "이름", Array("부서", "=", "개발", "OR", "부서", "=", "영업"))
'----------------------------------------------------------------------------
    
    Dim result As Range
    Dim i As Long, j As Long
    Dim conditionMet As Boolean
    Dim headerRow As Range
    Dim outputColIndices() As Long
    Dim outputColCount As Long
    Dim conditionResults() As Boolean
    Dim conditionCount As Long
    Dim logicOps() As String
    Dim condIdx As Long
    Dim outputCell As Range
    Dim processedConditions() As Variant
    
    ' 헤더 행 설정
    Set headerRow = targetTable.HeaderRowRange
    
    ' 출력 열 인덱스 설정
    If VarType(outputColumns) = vbString Then
        ReDim outputColIndices(0 To 0)
        outputColIndices(0) = GetColumnIndex(headerRow, CStr(outputColumns))
        If outputColIndices(0) = 0 Then
            MsgBox "출력 열 '" & outputColumns & "'을(를) 찾을 수 없습니다.", vbExclamation
            Set TblFindRng_MC = Nothing
            Exit Function
        End If
        outputColCount = 1
    Else
        ReDim outputColIndices(0 To UBound(outputColumns))
        For i = 0 To UBound(outputColumns)
            outputColIndices(i) = GetColumnIndex(headerRow, CStr(outputColumns(i)))
            If outputColIndices(i) = 0 Then
                MsgBox "출력 열 '" & outputColumns(i) & "'을(를) 찾을 수 없습니다.", vbExclamation
                Set TblFindRng_MC = Nothing
                Exit Function
            End If
        Next i
        outputColCount = UBound(outputColumns) + 1
    End If
    
    ' 배열로 전달된 경우 처리
    If UBound(conditions) = 0 And IsArray(conditions(0)) Then
        Dim tempArr() As Variant
        ReDim tempArr(LBound(conditions(0)) To UBound(conditions(0)))
        Dim k As Long
        For k = LBound(conditions(0)) To UBound(conditions(0))
            tempArr(k) = conditions(0)(k)
        Next k
        processedConditions = tempArr
    Else
        processedConditions = conditions
    End If
    
    ' 조건이 없는 경우
    If UBound(processedConditions) < 2 Then
        Set TblFindRng_MC = Nothing
        Exit Function
    End If
    
    ' 조건 파싱: 조건 개수와 논리 연산자 추출
    conditionCount = 0
    ReDim conditionResults(0 To 100)
    ReDim logicOps(0 To 100)
    
    j = 0
    Do While j <= UBound(processedConditions)
        If j + 2 <= UBound(processedConditions) Then
            conditionCount = conditionCount + 1
            
            ' 다음 요소가 논리 연산자인지 확인
            If j + 3 <= UBound(processedConditions) Then
                If UCase(Trim(CStr(processedConditions(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(processedConditions(j + 3)))) = "OR" Then
                    logicOps(conditionCount - 1) = UCase(Trim(CStr(processedConditions(j + 3))))
                    j = j + 4
                Else
                    logicOps(conditionCount - 1) = "AND"
                    j = j + 3
                End If
            Else
                logicOps(conditionCount - 1) = "AND"
                j = j + 3
            End If
        Else
            Exit Do
        End If
    Loop
    
    If conditionCount > 0 Then
        ReDim Preserve conditionResults(0 To conditionCount - 1)
        ReDim Preserve logicOps(0 To conditionCount - 1)
    Else
        Set TblFindRng_MC = Nothing
        Exit Function
    End If
    
    ' 각 행별로 검색
    For i = 1 To targetTable.ListRows.count
        
        ' 각 조건 평가
        j = 0
        condIdx = 0
        
        Do While condIdx < conditionCount
            Dim colName As String
            Dim operator As String
            Dim value As Variant
            Dim colIndex As Long
            Dim cellValue As Variant
            
            ' 조건 파라미터 추출
            colName = CStr(processedConditions(j))
            operator = CStr(processedConditions(j + 1))
            value = processedConditions(j + 2)
            
            ' 열 인덱스 찾기
            colIndex = GetColumnIndex(headerRow, colName)
            If colIndex = 0 Then
                conditionResults(condIdx) = False
            Else
                ' 셀 값 가져오기
                cellValue = targetTable.ListRows(i).Range.Cells(1, colIndex).value
                
                ' 조건 비교
                conditionResults(condIdx) = EvaluateCondition(cellValue, operator, value)
            End If
            
            ' 다음 조건으로 이동
            If j + 3 <= UBound(processedConditions) Then
                If UCase(Trim(CStr(processedConditions(j + 3)))) = "AND" Or _
                   UCase(Trim(CStr(processedConditions(j + 3)))) = "OR" Then
                    j = j + 4
                Else
                    j = j + 3
                End If
            Else
                j = j + 3
            End If
            
            condIdx = condIdx + 1
        Loop
        
        ' 논리 연산자를 사용하여 최종 결과 계산
        conditionMet = conditionResults(0)
        For condIdx = 1 To conditionCount - 1
            If logicOps(condIdx - 1) = "AND" Then
                conditionMet = conditionMet And conditionResults(condIdx)
            ElseIf logicOps(condIdx - 1) = "OR" Then
                conditionMet = conditionMet Or conditionResults(condIdx)
            End If
        Next condIdx
        
        ' 조건에 맞는 셀들을 결과에 추가
        If conditionMet Then
            For j = 0 To outputColCount - 1
                Set outputCell = targetTable.ListRows(i).Range.Cells(1, outputColIndices(j))
                If result Is Nothing Then
                    Set result = outputCell
                Else
                    Set result = Union(result, outputCell)
                End If
            Next j
        End If
    Next i
    
    Set TblFindRng_MC = result
End Function

Private Function EvaluateCondition(cellValue As Variant, operator As String, value As Variant) As Boolean
'----------------------------------------------------------------------------
' 목적: 단일 조건 평가
'----------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case LCase(Trim(operator))
        Case "="
            EvaluateCondition = (cellValue = value)
        Case "<>"
            EvaluateCondition = (cellValue <> value)
        Case ">"
            EvaluateCondition = (cellValue > value)
        Case ">="
            EvaluateCondition = (cellValue >= value)
        Case "<"
            EvaluateCondition = (cellValue < value)
        Case "<="
            EvaluateCondition = (cellValue <= value)
        Case "like"
            EvaluateCondition = (cellValue Like value)
        Case Else
            EvaluateCondition = False
    End Select
    
    If Err.Number <> 0 Then
        EvaluateCondition = False
        Err.Clear
    End If
End Function

Private Function GetColumnIndex(ByVal headerRow As Range, ByVal colName As String) As Long
'----------------------------------------------------------------------------
' 목적: 열 이름에 해당하는 열 인덱스를 반환
' 매개변수:
'   - headerRow: 테이블의 헤더 행
'   - colName: 찾을 열 이름
' 반환: 열 인덱스 (찾지 못한 경우 0)
'----------------------------------------------------------------------------
    Dim cell As Range
    For Each cell In headerRow.Cells
        If LCase(Trim(cell.value)) = LCase(Trim(colName)) Then
            GetColumnIndex = cell.Column - headerRow.Column + 1
            Exit Function
        End If
    Next cell
    GetColumnIndex = 0
End Function

'============================================================================
' 테스트 코드
'============================================================================

Sub Test_TblFindVals_MC_Single()
    ' 단일 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim result As Variant
    result = TblFindVals_MC(tbl, "이름", "나이", ">=", 30)
    
    If UBound(result) >= 0 Then
        Dim i As Long
        Debug.Print "=== 나이 >= 30인 직원 ==="
        For i = 0 To UBound(result)
            Debug.Print result(i)
        Next i
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindVals_MC_AND()
    ' AND 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim result As Variant
    result = TblFindVals_MC(tbl, "이름", "나이", ">=", 25, "AND", "부서", "=", "개발")
    
    If UBound(result) >= 0 Then
        Dim i As Long
        Debug.Print "=== 나이 >= 25 AND 부서 = 개발 ==="
        For i = 0 To UBound(result)
            Debug.Print result(i)
        Next i
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindVals_MC_OR()
    ' OR 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim result As Variant
    result = TblFindVals_MC(tbl, "이름", "부서", "=", "개발", "OR", "부서", "=", "영업")
    
    If UBound(result) >= 0 Then
        Dim i As Long
        Debug.Print "=== 부서 = 개발 OR 부서 = 영업 ==="
        For i = 0 To UBound(result)
            Debug.Print result(i)
        Next i
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindVals_MC_DynamicOR()
    ' 동적 OR 조건 테스트 - 배열 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim arrTables() As Variant
    arrTables = Array("개발", "영업")
    
    Dim strConditions As String
    Dim i As Long
    
    For i = 0 To UBound(arrTables)
        strConditions = strConditions & IIf(i = 0, "", ";OR;") & "부서;=;" & arrTables(i)
    Next
    
    Dim arr() As String
    arr = Split(strConditions, ";")
    
    Dim result As Variant
    result = TblFindVals_MC(tbl, "이름", arr)
    
    If IsArray(result) And UBound(result) >= 0 Then
        Debug.Print "=== 동적 OR 조건으로 찾은 이름들 ==="
        For i = 0 To UBound(result)
            Debug.Print result(i)
        Next i
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindRng_MC_Single()
    ' 단일 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim rng As Range
    Set rng = TblFindRng_MC(tbl, "이름", "나이", ">=", 30)
    
    If Not rng Is Nothing Then
        Debug.Print "찾은 범위: " & rng.Address
        rng.Select
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindRng_MC_AND()
    ' AND 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim rng As Range
    Set rng = TblFindRng_MC(tbl, Array("이름", "부서"), _
                            "나이", ">=", 25, "AND", _
                            "부서", "=", "개발")
    
    If Not rng Is Nothing Then
        Debug.Print "찾은 범위: " & rng.Address
        rng.Select
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindRng_MC_OR()
    ' OR 조건 테스트 - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim rng As Range
    Set rng = TblFindRng_MC(tbl, "이름", _
                            "부서", "=", "개발", "OR", _
                            "부서", "=", "영업")
    
    If Not rng Is Nothing Then
        Debug.Print "찾은 범위: " & rng.Address
        Debug.Print "찾은 셀 개수: " & rng.Cells.count
        rng.Select
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindRng_MC_DynamicOR()
    ' 동적 OR 조건 테스트 - 배열 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    ' 동적으로 조건 생성
    Dim arrTables() As Variant
    arrTables = Array("개발", "영업", "인사")
    
    Dim strConditions As String
    Dim i As Long
    
    ' ?? 중요: strConditions에 누적해야 합니다
    For i = 0 To UBound(arrTables)
        strConditions = strConditions & IIf(i = 0, "", ";OR;") & "부서;=;" & arrTables(i)
    Next
    
    Dim arr() As String
    arr = Split(strConditions, ";")
    
    ' 배열을 직접 전달
    Dim rng As Range
    Set rng = TblFindRng_MC(tbl, "이름", arr)
    
    If Not rng Is Nothing Then
        Debug.Print "동적 OR 조건 결과: " & rng.Address
        Debug.Print "찾은 셀 개수: " & rng.Cells.count
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

Sub Test_TblFindRng_MC_Mixed()
    ' 혼합 조건 테스트 (왼쪽→오른쪽 순차 평가) - 직접 인수 전달
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("T_직원")
    
    Dim rng As Range
    Set rng = TblFindRng_MC(tbl, Array("이름", "부서", "나이"), _
                            "나이", ">=", 25, "AND", _
                            "부서", "=", "개발", "OR", _
                            "직급", "=", "과장")
    
    If Not rng Is Nothing Then
        Debug.Print "찾은 범위: " & rng.Address
        rng.Select
    Else
        Debug.Print "조건에 맞는 데이터가 없습니다."
    End If
End Sub

