Attribute VB_Name = "tpl_MsSQL"
Option Explicit
Function conStr(Optional ByVal strType As String, _
                Optional ByVal strFile As String, _
                Optional ByVal strServer As String, _
                Optional ByVal strPort As String, _
                Optional ByVal strDB As String, _
                Optional ByVal strID As String, _
                Optional ByVal strPW As String, _
                Optional ByVal blnIncludeHeaders As Boolean = False) As String
    
    Dim extendedProps As String
    
    ' 빈 값이면 모듈 변수 사용
    If strType = "" Then strType = cfg_Type
    If strFile = "" Then strFile = cfg_File
    If strServer = "" Then strServer = cfg_Server
    If strPort = "" Then strPort = cfg_Port
    If strDB = "" Then strDB = cfg_DB
    If strID = "" Then strID = cfg_ID
    If strPW = "" Then strPW = cfg_PW
    
    Select Case strType
    Case "서버" 'Server
        conStr = "Driver={" & PublicODBC_Driver & "};Server=" & strServer & ";Port=" & strPort & ";Database=" & strDB & ";User=" & strID & ";Password=" & strPW & ";Option=3;"
    Case "엑셀" 'ExcelSheet
        ' Extended Properties 구성
        extendedProps = "Excel " & PublicExcel_Ver
        If blnIncludeHeaders Then
            extendedProps = extendedProps & ";HDR=Yes;IMEX=1"
        End If
        
        conStr = "Provider=Microsoft.ACE.OLEDB." & PublicExcel_Ver & ";Data Source=" & strFile & ";Extended Properties=""" & extendedProps & """"
    Case "엑세스" 'Access
        conStr = "Provider=Microsoft.ACE.OLEDB." & PublicAccess_Ver & ";Data Source=" & strFile & ";"
    End Select
    
End Function
Function ExecuteQueryArr(ByVal arrQuery As Variant, _
                        Optional ByVal strType As String, _
                        Optional ByVal strFile As String, _
                        Optional ByVal strServer As String, _
                        Optional ByVal strPort As String, _
                        Optional ByVal strDB As String, _
                        Optional ByVal strID As String, _
                        Optional ByVal strPW As String) As Boolean
    '<<< 다중 쿼리 실행 (엑셀 시트에 접속시 strFile까지 기록, 서버 데이터베이스 접속시 해당 값 제외 모두 기록) >>>
    Dim conn As Object
    Dim cmd As Object
    Dim connectionString As String
    Dim strQuery As String
    Dim strSheet As String
    Dim intEndRow As Long
    Dim ws As Worksheet
    Dim CellInsert As Range
    Dim i As Long
    Dim j As Long
    Dim strErrorLocation As String  ' 에러 발생 위치 추적
    Dim retryCount As Integer  ' 재시도 횟수
    Dim maxRetries As Integer  ' 최대 재시도 횟수
    
    maxRetries = 3  ' 최대 3번 재시도
    retryCount = 0
    
RetryStart:
    On Error GoTo ErrorHandler
    
    ' DB 타입 확인 (빈 값이면 모듈 변수 사용)
    If strType = "" Then strType = cfg_Type
    
    ' 연결 문자열 생성
    strErrorLocation = "연결 문자열 생성"
    Set conn = CreateObject("ADODB.Connection")
    connectionString = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW)
    
    
    '<< Excel 타입인 경우 마지막 행 찾기 시작 ==========================================================
    If strType = "엑셀" Then
        strErrorLocation = "워크시트 확인"
        
        ' 첫 번째 쿼리에서 시트명 추출
        strSheet = arrQuery(0)
        strSheet = Mid(strSheet, InStr(strSheet, "[") + 1, InStrRev(strSheet, "$") - InStr(strSheet, "[") - 1)
        
        ' 워크시트 존재 확인 (에러 발생 가능성 있음)
        Err.Clear
        On Error Resume Next
        Set ws = Worksheets(strSheet)
        On Error GoTo ErrorHandler
        
        If ws Is Nothing Or Err.Number <> 0 Then
            MsgBox "워크시트 '" & strSheet & "'를 찾을 수 없습니다.", vbCritical, "Worksheet Error"
            ExecuteQueryArr = False
            Exit Function
        End If
        
        ' 마지막 행 찾기
        strErrorLocation = "마지막 행 찾기"
        Set CellInsert = ws.Cells(ws.Rows.count, 1).End(xlUp)
        
        If CellInsert.value = "" Then
            intEndRow = CellInsert.Row - 1
        Else
            intEndRow = CellInsert.Row
        End If
    End If
    '<< Excel 타입인 경우 마지막 행 찾기 종료 ==========================================================
    
    ' DB 연결
    strErrorLocation = "데이터베이스 연결"
    Err.Clear
    conn.Open connectionString
    
    ' 쿼리 실행 준비
    strErrorLocation = "쿼리 실행 준비"
    Set cmd = CreateObject("ADODB.Command")
    
    With cmd
        .ActiveConnection = conn
        
        For i = 0 To UBound(arrQuery)
            strErrorLocation = "[" & i & "]번째 쿼리 실행"
            strQuery = arrQuery(i)
            
            ' Excel INSERT 쿼리 처리
            If strType = "엑셀" And InStr(strQuery, "INSERT") > 0 Then
                strQuery = Replace(strQuery, "$]", "$1:" & intEndRow + j & "]")
                j = j + 1
            End If
            
            ' 쿼리 실행 (에러 초기화 필수!)
            Err.Clear
            .CommandText = strQuery
            .Execute

        Next i
    End With
    
    ' 성공
    ExecuteQueryArr = True
    GoTo CleanUp
    
ErrorHandler:
    ' 재시도 가능 여부 확인
    retryCount = retryCount + 1
    
    ' 워크시트 누락 같은 명확한 오류는 재시도 안함
    If InStr(strErrorLocation, "워크시트") > 0 Then
        ExecuteQueryArr = False
        GoTo ShowError
    End If
    
    ' 최대 재시도 횟수 미만이면 재시도
    If retryCount <= maxRetries Then
        ' 연결 정리 후 재시도
        On Error Resume Next
        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
        End If
        Set cmd = Nothing
        Set conn = Nothing
        On Error GoTo 0
        
        Application.Wait Now + TimeValue("00:00:01")  ' 1초 대기
        GoTo RetryStart
    End If
    
ShowError:
    ExecuteQueryArr = False
    
    ' 에러 발생 위치와 상세 정보 출력
    MsgBox "오류 발생 위치: " & strErrorLocation & vbNewLine & vbNewLine & _
           "재시도 횟수: " & retryCount - 1 & "/" & maxRetries & vbNewLine & vbNewLine & _
           "Error Code: " & Err.Number & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & vbNewLine & _
           "Query: " & strQuery, vbCritical, "ExecuteQueryArr Error"
    Debug.Print "Error: " & strQuery
    
CleanUp:
    ' 연결 종료
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    On Error GoTo 0
    
    ' 메모리 정리
    Erase arrQuery
    Set cmd = Nothing
    Set conn = Nothing
    Set ws = Nothing
    Set CellInsert = Nothing
End Function
Public Sub SelectQuery( _
    ByVal celTarget As Range, _
    ByVal arrQuery As Variant, _
    Optional ByVal strType As String, _
    Optional ByVal strFile As String, _
    Optional ByVal strServer As String, _
    Optional ByVal strPort As String, _
    Optional ByVal strDB As String, _
    Optional ByVal strID As String, _
    Optional ByVal strPW As String, _
    Optional ByVal BlnTranspose As Boolean = False, _
    Optional ByVal intMoveCells As Integer = 0, _
    Optional ByVal blnIncludeHeaders As Boolean = False _
)

    Dim conn As Object
    Dim rs As Object
    Dim connectionString As String
    Dim strQuery As String
    Dim i As Long, j As Long
    Dim targetCell As Range
    Dim originalSheet As Worksheet
    Dim retryCount As Integer
    Dim maxRetries As Integer
    
    maxRetries = 3  ' 최대 3번 재시도
    retryCount = 0
    
    Set originalSheet = ActiveSheet

RetryStart:
    On Error GoTo ErrorHandler

    ' ===== Connection =====
    Set conn = CreateObject("ADODB.Connection")
    connectionString = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW, blnIncludeHeaders)
    conn.Open connectionString

    ' ===== Recordset =====
    Set rs = CreateObject("ADODB.Recordset")

    For i = 0 To UBound(arrQuery)
        strQuery = CStr(arrQuery(i))
        ' 기본 ForwardOnly / ReadOnly
        rs.Open strQuery, conn

        If Not rs.EOF Then
            If Not celTarget Is Nothing Then

                Set targetCell = celTarget

                ' --- 두 번째 쿼리부터 위치 조정 ---
                If i > 0 Then
                    If BlnTranspose Then
                        If intMoveCells > 0 Then
                            Set targetCell = celTarget.Offset(, i * intMoveCells)
                        Else
                            With celTarget.Worksheet
                                Set targetCell = .Cells(celTarget.Row, .Columns.count).End(xlToLeft)
                                If targetCell.Value2 <> "" Then
                                    Set targetCell = targetCell.Offset(, 1)
                                End If
                            End With
                        End If
                    Else
                        If intMoveCells > 0 Then
                            Set targetCell = celTarget.Offset(i * intMoveCells, 0)
                        Else
                            With celTarget.Worksheet
                                Set targetCell = .Cells(.Rows.count, celTarget.Column).End(xlUp)
                                If targetCell.Value2 <> "" Then
                                    Set targetCell = targetCell.Offset(1, 0)
                                End If
                            End With
                        End If
                    End If
                End If

                ' 시트 활성화
                If Not (targetCell.Worksheet Is ActiveSheet) Then
                    targetCell.Worksheet.Activate
                End If

                ' ===== 헤더 출력 =====
                If blnIncludeHeaders Then
                    If BlnTranspose Then
                        For j = 0 To rs.fields.count - 1
                            targetCell.Offset(j, 0).Value2 = rs.fields(j).Name
                        Next j
                        Set targetCell = targetCell.Offset(0, 1)
                    Else
                        For j = 0 To rs.fields.count - 1
                            targetCell.Offset(0, j).Value2 = rs.fields(j).Name
                        Next j
                        Set targetCell = targetCell.Offset(1, 0)
                    End If
                End If

                ' ===== 핵심: CopyFromRecordset =====
                targetCell.CopyFromRecordset rs

            End If
        End If

        rs.Close
    Next i

Exit Sub

ErrorHandler:
    retryCount = retryCount + 1
    
    ' 최대 재시도 횟수 미만이면 재시도
    If retryCount <= maxRetries Then
        ' 연결 정리 후 재시도
        On Error Resume Next
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
            Set conn = Nothing
        End If
        On Error GoTo 0
        
        Application.Wait Now + TimeValue("00:00:01")  ' 1초 대기
        GoTo RetryStart
    End If
    
    ' 최종 에러 처리
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Query: " & strQuery
    Debug.Print "재시도 횟수: " & retryCount - 1 & "/" & maxRetries
    Call HandleError("SelectQuery")
End Sub
Function SelectQueryArr(ByVal arrQuery As Variant, _
                        Optional ByVal strType As String, _
                        Optional ByVal strFile As String, _
                        Optional ByVal strServer As String, _
                        Optional ByVal strPort As String, _
                        Optional ByVal strDB As String, _
                        Optional ByVal strID As String, _
                        Optional ByVal strPW As String, _
                        Optional ByVal blnIncludeHeaders As Boolean = False, _
                        Optional ByVal BlnTranspose As Boolean = False) As Variant
                                                
    ' 여러 개의 Query 쿼리를 실행하고 각 결과를 배열로 반환
    ' 반환값: 배열의 배열 (각 쿼리 결과가 별도 배열로 저장)
    Dim conn As Object
    Dim rs As Object
    Dim connectionString As String
    Dim resultCollection() As Variant
    Dim queryIndex As Long
    Dim i As Long, ri As Long, ci As Long
    Dim arrTemp As Variant, arrReTemp As Variant
    Dim intRow As Long, intColumn As Long
    Dim headerOffset As Long
    Dim retryCount As Integer
    Dim maxRetries As Integer
    
    maxRetries = 3  ' 최대 3번 재시도
    retryCount = 0
    
RetryStart:
    ' 연결 객체 생성
    Set conn = CreateObject("ADODB.Connection")
    connectionString = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW, blnIncludeHeaders)
    
    ' DB 연결
    On Error GoTo ConnectionError
    conn.Open connectionString
    On Error GoTo 0
    
    ' 결과 배열 초기화
    ReDim resultCollection(LBound(arrQuery) To UBound(arrQuery))
    
    ' 각 쿼리 실행
    For queryIndex = LBound(arrQuery) To UBound(arrQuery)
        Set rs = CreateObject("ADODB.Recordset")
        
        On Error GoTo ErrorPass
        rs.Open arrQuery(queryIndex), conn
        On Error GoTo 0
        
        ' 결과가 있는 경우
        If Not rs.EOF Then
            arrTemp = rs.GetRows()
            
            ' 마지막 유효 행 찾기 (모든 열이 Null인 행 제거)
            For ri = UBound(arrTemp, 2) To 0 Step -1
                For ci = 0 To UBound(arrTemp, 1)
                    If Not IsNull(arrTemp(ci, ri)) Then GoTo FoundLastRow
                Next ci
            Next ri
            
FoundLastRow:
            intRow = ri
            ReDim Preserve arrTemp(UBound(arrTemp, 1), intRow)
            
            ' 마지막 유효 열 찾기 (모든 행이 Null인 열 제거)
            For ci = UBound(arrTemp, 1) To 0 Step -1
                For ri = 0 To UBound(arrTemp, 2)
                    If Not IsNull(arrTemp(ci, ri)) Then GoTo FoundLastColumn
                Next ri
            Next ci
            
FoundLastColumn:
            intColumn = ci
            
            ' 헤더 포함 여부에 따른 오프셋 설정
            headerOffset = IIf(blnIncludeHeaders, 1, 0)
            
            ' ===== Transpose 여부에 따라 결과 배열 구성 =====
            If BlnTranspose Then
                ' Transpose: (행, 열) → arrReTemp(row, col) 형태로 재구성
                ' GetRows()는 (col, row) 구조이므로 행열 전환
                ReDim arrReTemp(intRow + headerOffset, intColumn)
                
                ' 헤더 출력 (첫 번째 열에 필드명)
                If blnIncludeHeaders Then
                    For ci = 0 To intColumn
                        arrReTemp(0, ci) = rs.fields(ci).Name
                    Next ci
                End If
                
                ' 데이터 복사 (행열 전환)
                For ri = 0 To intRow
                    For ci = 0 To intColumn
                        arrReTemp(ri + headerOffset, ci) = arrTemp(ci, ri)
                    Next ci
                Next ri
                
            Else
                ' 기존 방식: (col, row) 구조 유지
                ReDim arrReTemp(intColumn, intRow + headerOffset)
                
                ' 필드명 추가
                If blnIncludeHeaders Then
                    For ci = 0 To intColumn
                        arrReTemp(ci, 0) = rs.fields(ci).Name
                    Next ci
                End If
                
                ' 데이터 복사
                For ri = 0 To intRow
                    For ci = 0 To intColumn
                        arrReTemp(ci, ri + headerOffset) = arrTemp(ci, ri)
                    Next ci
                Next ri
            End If
            
            resultCollection(queryIndex) = arrReTemp
        Else
            ' 결과가 없는 경우 빈 배열
            resultCollection(queryIndex) = Array()
        End If
        
        rs.Close
        Set rs = Nothing
        
        GoTo NextQuery
        
ErrorPass:
        Debug.Print "Error in Query " & queryIndex & ": " & arrQuery(queryIndex)
        MsgBox "Query 구문오류! (쿼리 인덱스: " & queryIndex & ")", vbCritical
        resultCollection(queryIndex) = Array()
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        
NextQuery:
    Next queryIndex
    
    ' 연결 종료
    conn.Close
    Set conn = Nothing
    
    ' 결과 반환
    SelectQueryArr = resultCollection
    Exit Function

ConnectionError:
    retryCount = retryCount + 1
    
    ' 최대 재시도 횟수 미만이면 재시도
    If retryCount <= maxRetries Then
        ' 연결 정리 후 재시도
        On Error Resume Next
        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
            Set conn = Nothing
        End If
        On Error GoTo 0
        
        Application.Wait Now + TimeValue("00:00:01")  ' 1초 대기
        GoTo RetryStart
    End If
    
    ' 최종 에러 처리
    Debug.Print "Connection Error after " & retryCount - 1 & " retries"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    SelectQueryArr = Array()
End Function
Sub DelExcelRecQuery(ByVal strSheetName As String, _
                                            ByVal strID As String)
    
    Dim sht As Worksheet
    Dim f, rng As Range
    Dim intRow  As Long
    
    Set sht = ThisWorkbook.Worksheets(strSheetName)
    Set rng = sht.Range("A:A")
    
    Set f = rng.Find(strID, , xlValues, xlWhole)
    
    If Not f Is Nothing Then f.EntireRow.Delete
    
End Sub
Function GetDbTables(ByVal strType As String, ByVal strFilePath As String) As Variant
    ' strType: "엑셀" 또는 "엑세스"
    ' strFilePath: 파일 경로
    
    Dim conn As Object
    Dim rs As Object
    Dim connectionString As String
    Dim arrResult() As String
    Dim i As Long
    
    Set conn = CreateObject("ADODB.Connection")
    
    ' 연결 문자열 설정
    Select Case strType
        Case "엑셀"
            connectionString = "Provider=Microsoft.ACE.OLEDB." & PublicExcel_Ver & ";Data Source=" & strFilePath & ";Extended Properties=Excel " & PublicExcel_Ver & ";"
        Case "엑세스"
            connectionString = "Provider=Microsoft.ACE.OLEDB." & PublicAccess_Ver & ";Data Source=" & strFilePath & ";"
        Case Else
            MsgBox "지원하지 않는 타입입니다. '엑셀' 또는 '엑세스'를 입력하세요.", vbCritical
            Exit Function
    End Select
    
    conn.Open connectionString
    
    ' 테이블/시트 스키마 정보 가져오기
    Set rs = conn.OpenSchema(20) ' adSchemaTables = 20
    
    i = 0
    Do While Not rs.EOF
        Select Case strType
            Case "엑셀"
                ' 시트명은 "Sheet1$" 형태로 반환되므로 $ 포함된 것만 추출
                If InStr(rs.fields("TABLE_NAME").value, "$") > 0 And _
                   InStr(rs.fields("TABLE_NAME").value, "'") = 0 Then ' 명명된 범위 제외
                    ReDim Preserve arrResult(i)
                    arrResult(i) = Replace(rs.fields("TABLE_NAME").value, "$", "")
                    i = i + 1
                End If
                
            Case "엑세스"
                ' 시스템 테이블 제외
                If rs.fields("TABLE_TYPE").value = "TABLE" And _
                   Left(rs.fields("TABLE_NAME").value, 4) <> "MSys" Then
                    ReDim Preserve arrResult(i)
                    arrResult(i) = rs.fields("TABLE_NAME").value
                    i = i + 1
                End If
        End Select
        
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
    
    GetDbTables = arrResult
End Function
Function GetFieldInfo(ByVal strType As String, _
                     ByVal strTable As String, _
                     Optional ByVal strFile As String, _
                     Optional ByVal strServer As String, _
                     Optional ByVal strPort As String, _
                     Optional ByVal strDB As String, _
                     Optional ByVal strID As String, _
                     Optional ByVal strPW As String) As Variant

    '[0] 순번 (1, 2, 3, ...)
    '[1] 필드명
    '[2] 데이터타입
    '[3] 길이
    '[4] NULL허용(YES / NO)/엑셀은 빈값 반환
    '[5] 데이터수
    
    Dim conn As Object
    Dim rs As Object
    Dim rsCount As Object
    Dim fld As Object
    Dim connectionString As String
    Dim arrResult() As Variant
    Dim i As Long
    Dim strQuery As String
    Dim strCountQuery As String
    Dim dictDataCount As Object  ' Dictionary로 데이터수 저장
    
    Set conn = CreateObject("ADODB.Connection")
    Set dictDataCount = CreateObject("Scripting.Dictionary")
    
    connectionString = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW)
    conn.Open connectionString
    
    Select Case strType
        Case "서버"
            ' MySQL - information_schema 사용
            strQuery = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE " & _
                    "FROM information_schema.COLUMNS " & _
                    "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = '" & strTable & "' " & _
                    "ORDER BY ORDINAL_POSITION;"
            
            Set rs = conn.Execute(strQuery)
            
            ' 먼저 모든 필드명 수집
            Dim arrFieldNames() As String
            Dim fieldCount As Long
            fieldCount = 0
            
            Do While Not rs.EOF
                ReDim Preserve arrFieldNames(fieldCount)
                arrFieldNames(fieldCount) = rs.fields("COLUMN_NAME").value
                fieldCount = fieldCount + 1
                rs.MoveNext
            Loop
            rs.Close
            
            ' 한 번의 쿼리로 모든 필드의 데이터 수 가져오기
            If fieldCount > 0 Then
                strCountQuery = "SELECT "
                For i = 0 To fieldCount - 1
                    If i > 0 Then strCountQuery = strCountQuery & ", "
                    strCountQuery = strCountQuery & "COUNT(`" & arrFieldNames(i) & "`) AS cnt" & i
                Next i
                strCountQuery = strCountQuery & " FROM `" & strTable & "`"
                
                On Error Resume Next
                Set rsCount = conn.Execute(strCountQuery)
                
                If Err.Number = 0 And Not rsCount.EOF Then
                    For i = 0 To fieldCount - 1
                        dictDataCount.Add arrFieldNames(i), rsCount.fields("cnt" & i).value
                    Next i
                End If
                
                If Not rsCount Is Nothing Then rsCount.Close
                On Error GoTo 0
            End If
            
            ' 다시 필드 정보 조회
            Set rs = conn.Execute(strQuery)
            
            i = 0
            Do While Not rs.EOF
                ReDim Preserve arrResult(5, i)  ' 5로 증가 (순번 추가)
                arrResult(0, i) = i + 1  '번 추가
                arrResult(1, i) = rs.fields("COLUMN_NAME").value
                arrResult(2, i) = rs.fields("DATA_TYPE").value
                arrResult(3, i) = IIf(IsNull(rs.fields("CHARACTER_MAXIMUM_LENGTH").value), "", rs.fields("CHARACTER_MAXIMUM_LENGTH").value)
                arrResult(4, i) = rs.fields("IS_NULLABLE").value
                
                ' Dictionary에서 데이터수 가져오기
                If dictDataCount.Exists(rs.fields("COLUMN_NAME").value) Then
                    arrResult(5, i) = dictDataCount(rs.fields("COLUMN_NAME").value)
                Else
                    arrResult(5, i) = 0
                End If
                
                i = i + 1
                rs.MoveNext
            Loop
            rs.Close
            
        Case "엑셀"
            ' 엑셀 - 실제 데이터를 조회하여 필드명과 타입 가져오기
            If Right(strTable, 1) <> "$" Then strTable = strTable & "$"
            
            strQuery = "SELECT * FROM [" & strTable & "]"
            Set rs = conn.Execute(strQuery)
            
            ' 유효한 필드명 수집
            ReDim arrFieldNames(rs.fields.count - 1)
            fieldCount = 0
            
            For Each fld In rs.fields
                If Not IsAutoGeneratedFieldName(fld.Name) And Trim(fld.Name) <> "" Then
                    arrFieldNames(fieldCount) = fld.Name
                    fieldCount = fieldCount + 1
                End If
            Next fld
            
            ' 필드명 배열 크기 조정
            If fieldCount > 0 Then
                ReDim Preserve arrFieldNames(fieldCount - 1)
                
                ' 한 번의 쿼리로 모든 필드의 데이터 수 가져오기
                strCountQuery = "SELECT "
                For i = 0 To fieldCount - 1
                    If i > 0 Then strCountQuery = strCountQuery & ", "
                    strCountQuery = strCountQuery & "COUNT([" & arrFieldNames(i) & "]) AS cnt" & i
                Next i
                strCountQuery = strCountQuery & " FROM [" & strTable & "]"
                
                On Error Resume Next
                Set rsCount = conn.Execute(strCountQuery)
                
                If Err.Number = 0 And Not rsCount.EOF Then
                    For i = 0 To fieldCount - 1
                        dictDataCount.Add arrFieldNames(i), rsCount.fields("cnt" & i).value
                    Next i
                End If
                
                If Not rsCount Is Nothing Then rsCount.Close
                On Error GoTo 0
            End If
            
            ' 필드 정보 수집
            i = 0
            For Each fld In rs.fields
                If Not IsAutoGeneratedFieldName(fld.Name) And Trim(fld.Name) <> "" Then
                    ReDim Preserve arrResult(5, i)  ' 5로 변경 (서버/엑세스와 동일하게)
                    arrResult(0, i) = ""  ' 순번은 빈값 (엑셀은 순번 정보 없음)
                    arrResult(1, i) = fld.Name
                    arrResult(2, i) = GetDataTypeName(fld.Type)
                    arrResult(3, i) = IIf(fld.DefinedSize > 0, fld.DefinedSize, "")
                    arrResult(4, i) = ""  ' NULL허용 정보 없음 (엑셀)
                    
                    ' Dictionary에서 데이터수 가져오기
                    If dictDataCount.Exists(fld.Name) Then
                        arrResult(5, i) = dictDataCount(fld.Name)
                    Else
                        arrResult(5, i) = 0
                    End If
                    
                    i = i + 1
                End If
            Next fld
            rs.Close
            
        Case "엑세스"
            ' 엑세스 - OpenSchema 사용
            Set rs = conn.OpenSchema(4, Array(Empty, Empty, strTable))
            
            ' 필드명 수집
            fieldCount = 0
            Do While Not rs.EOF
                ReDim Preserve arrFieldNames(fieldCount)
                arrFieldNames(fieldCount) = rs.fields("COLUMN_NAME").value
                fieldCount = fieldCount + 1
                rs.MoveNext
            Loop
            
            ' 한 번의 쿼리로 모든 필드의 데이터 수 가져오기
            If fieldCount > 0 Then
                strCountQuery = "SELECT "
                For i = 0 To fieldCount - 1
                    If i > 0 Then strCountQuery = strCountQuery & ", "
                    strCountQuery = strCountQuery & "COUNT([" & arrFieldNames(i) & "]) AS cnt" & i
                Next i
                strCountQuery = strCountQuery & " FROM [" & strTable & "]"
                
                On Error Resume Next
                Set rsCount = conn.Execute(strCountQuery)
                
                If Err.Number = 0 And Not rsCount.EOF Then
                    For i = 0 To fieldCount - 1
                        dictDataCount.Add arrFieldNames(i), rsCount.fields("cnt" & i).value
                    Next i
                End If
                
                If Not rsCount Is Nothing Then rsCount.Close
                On Error GoTo 0
            End If
            
            ' 필드 정보 수집
            rs.Close
            Set rs = conn.OpenSchema(4, Array(Empty, Empty, strTable))
            
            Dim tempArr() As Variant
            Dim j As Long, k As Long, L As Long
            Dim temp As Variant
            
            i = 0
            Do While Not rs.EOF
                ReDim Preserve tempArr(6, i)  ' 6으로 증가 (순번 추가)
                tempArr(0, i) = rs.fields("ORDINAL_POSITION").value  ' 실제 순번 저장
                tempArr(1, i) = rs.fields("COLUMN_NAME").value
                tempArr(2, i) = GetDataTypeName(rs.fields("DATA_TYPE").value)
                tempArr(3, i) = IIf(IsNull(rs.fields("CHARACTER_MAXIMUM_LENGTH").value), "", rs.fields("CHARACTER_MAXIMUM_LENGTH").value)
                tempArr(4, i) = IIf(IsNull(rs.fields("IS_NULLABLE").value), "", rs.fields("IS_NULLABLE").value)
                tempArr(5, i) = rs.fields("ORDINAL_POSITION").value ' 정렬용
                
                ' Dictionary에서 데이터수 가져오기
                If dictDataCount.Exists(rs.fields("COLUMN_NAME").value) Then
                    tempArr(6, i) = dictDataCount(rs.fields("COLUMN_NAME").value)
                Else
                    tempArr(6, i) = 0
                End If
                
                i = i + 1
                rs.MoveNext
            Loop
            rs.Close
            
            ' ORDINAL_POSITION으로 정렬 (버블 정렬)
            If i > 0 Then
                For j = 0 To i - 2
                    For k = j + 1 To i - 1
                        If tempArr(5, j) > tempArr(5, k) Then
                            For L = 0 To 6  ' 6으로 변경
                                temp = tempArr(L, j)
                                tempArr(L, j) = tempArr(L, k)
                                tempArr(L, k) = temp
                            Next L
                        End If
                    Next k
                Next j
                
                ' 결과 배열에 복사 (정렬용 ORDINAL_POSITION 제외)
                ReDim arrResult(5, i - 1)  ' 5로 증가
                For j = 0 To i - 1
                    arrResult(0, j) = tempArr(0, j)  ' 순번
                    arrResult(1, j) = tempArr(1, j)  ' 필드명
                    arrResult(2, j) = tempArr(2, j)  ' 데이터타입
                    arrResult(3, j) = tempArr(3, j)  ' 길이
                    arrResult(4, j) = tempArr(4, j)  ' NULL허용
                    arrResult(5, j) = tempArr(6, j)  ' 데이터수
                Next j
            End If
    End Select
    
    conn.Close
    Set rs = Nothing
    Set rsCount = Nothing
    Set conn = Nothing
    Set dictDataCount = Nothing
    
    ' 배열 전치 (ListBox에 바로 할당할 수 있도록)
    Dim finalResult() As Variant
    Dim m As Long, n As Long
    Dim rowCount As Long, colCount As Long
    
    On Error Resume Next
    rowCount = UBound(arrResult, 2) - LBound(arrResult, 2) + 1
    colCount = UBound(arrResult, 1) - LBound(arrResult, 1) + 1
    On Error GoTo 0
    
    If rowCount > 0 And colCount > 0 Then
        ReDim finalResult(0 To rowCount - 1, 0 To colCount - 1)
        
        For m = LBound(arrResult, 1) To UBound(arrResult, 1)
            For n = LBound(arrResult, 2) To UBound(arrResult, 2)
                finalResult(n, m) = arrResult(m, n)
            Next n
        Next m
        
        GetFieldInfo = finalResult
    Else
        GetFieldInfo = Array()
    End If
End Function
Function IsAutoGeneratedFieldName(fieldName As String) As Boolean '자동 생성된 필드명인지 확인하는 함수
    IsAutoGeneratedFieldName = (fieldName Like "F[0-9]" Or fieldName Like "F[0-9][0-9]" Or fieldName Like "F[0-9][0-9][0-9]")
End Function
' 데이터 타입 코드를 이름으로 변환하는 헬퍼 함수
Function GetDataTypeName(ByVal intTypeCode As Integer) As String
    Select Case intTypeCode
        Case 2: GetDataTypeName = "SmallInt"           '작은 정수 (-32,768 ~ 32,767)
        Case 3: GetDataTypeName = "Integer"            '정수 (-2,147,483,648 ~ 2,147,483,647)
        Case 4: GetDataTypeName = "Single"             '단정밀도 실수 (소수점 7자리)
        Case 5: GetDataTypeName = "Double"             '배정밀도 실수 (소수점 15자리)
        Case 6: GetDataTypeName = "Currency"           '통화 (고정소수점, 4자리)
        Case 7: GetDataTypeName = "Date"               '날짜/시간
        Case 11: GetDataTypeName = "Boolean"           '참/거짓 (True/False)
        Case 17: GetDataTypeName = "TinyInt"           '아주 작은 정수 (0 ~ 255)
        Case 20: GetDataTypeName = "BigInt"            '큰 정수 (64비트)
        Case 72: GetDataTypeName = "GUID"              '전역 고유 식별자 (UUID)
        Case 128: GetDataTypeName = "Binary"           '고정길이 이진 데이터
        Case 129: GetDataTypeName = "Char"             '고정길이 문자열
        Case 130: GetDataTypeName = "VarChar"          '가변길이 문자열
        Case 131: GetDataTypeName = "Numeric"          '고정 정밀도 숫자
        Case 200: GetDataTypeName = "VarChar"          '짧은 텍스트 (가변길이)
        Case 201: GetDataTypeName = "LongText"         '긴 텍스트 (메모)
        Case 202: GetDataTypeName = "VarWChar"         '유니코드 가변길이 문자열
        Case 203: GetDataTypeName = "LongVarWChar"     '긴 유니코드 텍스트
        Case 204: GetDataTypeName = "VarBinary"        '가변길이 이진 데이터
        Case 205: GetDataTypeName = "LongVarBinary"    '긴 이진 데이터 (이미지, 파일 등)
        Case Else: GetDataTypeName = "Unknown (" & intTypeCode & ")"  '알 수 없는 타입
    End Select
End Function
Function GetDataTypeCode(ByVal strTypeName As String) As Integer
    Dim strType As String
    
    ' 대소문자 구분 없이 처리하기 위해 대문자로 변환
    strType = UCase(Trim(strTypeName))
    
    Select Case strType
        Case "SMALLINT": GetDataTypeCode = 2           '작은 정수 (-32,768 ~ 32,767)
        Case "INTEGER": GetDataTypeCode = 3            '정수 (-2,147,483,648 ~ 2,147,483,647)
        Case "SINGLE": GetDataTypeCode = 4             '단정밀도 실수 (소수점 7자리)
        Case "DOUBLE": GetDataTypeCode = 5             '배정밀도 실수 (소수점 15자리)
        Case "CURRENCY": GetDataTypeCode = 6           '통화 (고정소수점, 4자리)
        Case "DATE": GetDataTypeCode = 7               '날짜/시간
        Case "BOOLEAN": GetDataTypeCode = 11           '참/거짓 (True/False)
        Case "TINYINT": GetDataTypeCode = 17           '아주 작은 정수 (0 ~ 255)
        Case "BIGINT": GetDataTypeCode = 20            '큰 정수 (64비트)
        Case "GUID": GetDataTypeCode = 72              '전역 고유 식별자 (UUID)
        Case "BINARY": GetDataTypeCode = 128           '고정길이 이진 데이터
        Case "CHAR": GetDataTypeCode = 129             '고정길이 문자열
        Case "VARCHAR": GetDataTypeCode = 130          '가변길이 문자열 또는 202
        Case "NUMERIC": GetDataTypeCode = 131          '고정 정밀도 숫자
        Case "LONGTEXT": GetDataTypeCode = 201         '긴 텍스트 (메모)
        Case "VARWCHAR": GetDataTypeCode = 202         '유니코드 가변길이 문자열
        Case "LONGVARWCHAR": GetDataTypeCode = 203     '긴 유니코드 텍스트
        Case "VARBINARY": GetDataTypeCode = 204        '가변길이 이진 데이터
        Case "LONGVARBINARY": GetDataTypeCode = 205    '긴 이진 데이터 (이미지, 파일 등)
        
        ' 별칭 처리
        Case "TEXT": GetDataTypeCode = 202             '텍스트 별칭
        Case "MEMO": GetDataTypeCode = 203             '메모 별칭
        Case "YESNO": GetDataTypeCode = 11             'Yes/No 별칭
        Case "LONG": GetDataTypeCode = 3               'Long 별칭
        Case "INT": GetDataTypeCode = 3                'Int 별칭
        Case "DATETIME": GetDataTypeCode = 7           'DateTime 별칭
        
        Case Else: GetDataTypeCode = -1                '알 수 없는 타입
    End Select
End Function
Sub AutoIDs(ByVal rng As Range)
    
    Dim arrIDs() As Long
    Dim r As Range
    Dim i, intID, intMaxID As Long
    
    intMaxID = Application.WorksheetFunction.Max(rng)
    intID = intMaxID + 1
    
    For Each r In rng
        ReDim Preserve arrIDs(i)
        If r = "" Then
            arrIDs(i) = intID
            intID = intID + 1
        Else
            arrIDs(i) = r
        End If
        i = i + 1
    Next
    
    rng = Application.Transpose(arrIDs)
    
End Sub


