Attribute VB_Name = "am_DB"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_DB                                                  │
' │  역할 : DB 연결, 쿼리 실행, 스키마 조회, 타입 처리     │
' └─────────────────────────────────────────────────────────┘

' ── 1. 상수 ──────────────────────────────────────────────────
Public Const DB_ODBC_DRIVER As String = "MySQL ODBC 8.2 Unicode Driver"
Public Const DB_EXCEL_VER   As String = "12.0"
Public Const DB_ACCESS_VER  As String = "12.0"

Public Const DB_TYPE_SERVER As String = "서버"
Public Const DB_TYPE_EXCEL  As String = "엑셀"
Public Const DB_TYPE_ACCESS As String = "엑세스"

' ══════════════════════════════════════════════════════════
'  연결 문자열
' ══════════════════════════════════════════════════════════

' 목적   : DB 연결 문자열 생성
' 인수   : strType   - DB 종류 ("서버" / "엑셀" / "엑세스")
'          strFile   - 파일 경로 (엑셀/엑세스)
'          strServer - 서버 주소 (서버)
'          strPort   - 포트 번호 (서버)
'          strDB     - DB 명 (서버)
'          strID     - 접속 ID (서버)
'          strPW     - 접속 PW (서버)
'          blnHeader - True: 헤더 있음 (엑셀 전용)
' 반환   : String - 연결 문자열
' 예시   : conStr("엑세스", "C:\DB\main.accdb")
Public Function conStr(ByVal strType As String, _
                       Optional ByVal strFile As String, _
                       Optional ByVal strServer As String, _
                       Optional ByVal strPort As String, _
                       Optional ByVal strDB As String, _
                       Optional ByVal strID As String, _
                       Optional ByVal strPW As String, _
                       Optional ByVal blnHeader As Boolean = False) As String

    Dim strExtProps As String

    Select Case strType
        Case DB_TYPE_SERVER
            conStr = "Driver={" & DB_ODBC_DRIVER & "};" & _
                     "Server=" & strServer & ";" & _
                     "Port=" & strPort & ";" & _
                     "Database=" & strDB & ";" & _
                     "User=" & strID & ";" & _
                     "Password=" & strPW & ";" & _
                     "Option=3;"

        Case DB_TYPE_EXCEL
            strExtProps = "Excel " & DB_EXCEL_VER
            If blnHeader Then strExtProps = strExtProps & ";HDR=Yes;IMEX=1"
            conStr = "Provider=Microsoft.ACE.OLEDB." & DB_EXCEL_VER & ";" & _
                     "Data Source=" & strFile & ";" & _
                     "Extended Properties=""" & strExtProps & """"

        Case DB_TYPE_ACCESS
            conStr = "Provider=Microsoft.ACE.OLEDB." & DB_ACCESS_VER & ";" & _
                     "Data Source=" & strFile & ";"

    End Select

End Function

' ══════════════════════════════════════════════════════════
'  쿼리 실행
' ══════════════════════════════════════════════════════════

' 목적   : 변경 쿼리 배열 실행 (INSERT / UPDATE / DELETE)
' 인수   : arrQuery  - 실행할 쿼리 배열
'          strType   - DB 종류
'          strFile   - 파일 경로
'          strServer - 서버 주소
'          strPort   - 포트 번호
'          strDB     - DB 명
'          strID     - 접속 ID
'          strPW     - 접속 PW
' 반환   : Boolean - True: 성공 / False: 실패
' 예시   : ExecuteQueryArr(Array("DELETE FROM T1 WHERE ID=1"), "엑세스", "C:\DB\main.accdb")
Public Function ExecuteQueryArr(ByVal arrQuery As Variant, _
                                ByVal strType As String, _
                                Optional ByVal strFile As String, _
                                Optional ByVal strServer As String, _
                                Optional ByVal strPort As String, _
                                Optional ByVal strDB As String, _
                                Optional ByVal strID As String, _
                                Optional ByVal strPW As String) As Boolean

    Dim objConn     As Object
    Dim objCmd      As Object
    Dim strConStr   As String
    Dim strQuery    As String
    Dim strErrLoc   As String
    Dim intRetry    As Integer
    Dim intMaxRetry As Integer
    Dim ws          As Worksheet
    Dim rngInsert   As Range
    Dim lngEndRow   As Long
    Dim strSheet    As String
    Dim i           As Long
    Dim j           As Long

    intMaxRetry = 3

RetryStart:
    On Error GoTo ErrHandler

    ' ── 1. 연결 문자열 생성 ───────────────────────────────────
    strErrLoc = "연결 문자열 생성"
    Set objConn = CreateObject("ADODB.Connection")
    strConStr = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW)

    ' ── 2. 엑셀 타입 시트명 확인 ────────────────────────────
    If strType = DB_TYPE_EXCEL Then
        strErrLoc = "워크시트 확인"
        strSheet = Mid(arrQuery(0), InStr(arrQuery(0), "[") + 1, _
                        InStrRev(arrQuery(0), "$") - InStr(arrQuery(0), "[") - 1)

        On Error Resume Next
        Set ws = Worksheets(strSheet)
        On Error GoTo ErrHandler

        If ws Is Nothing Then
            MsgBox "워크시트 '" & strSheet & "'을(를) 찾을 수 없습니다.", _
                   vbCritical, am_Core.AM_NAME
            ExecuteQueryArr = False
            Exit Function
        End If

        strErrLoc = "마지막 행 찾기"
        Set rngInsert = ws.Cells(ws.Rows.Count, 1).End(xlUp)
        lngEndRow = IIf(rngInsert.Value = "", rngInsert.Row - 1, rngInsert.Row)
    End If

    ' ── 3. DB 연결 ───────────────────────────────────────────
    strErrLoc = "DB 연결"
    objConn.Open strConStr

    ' ── 4. 쿼리 실행 ─────────────────────────────────────────
    strErrLoc = "쿼리 실행 준비"
    Set objCmd = CreateObject("ADODB.Command")

    With objCmd
        .ActiveConnection = objConn
        For i = LBound(arrQuery) To UBound(arrQuery)
            strErrLoc = i & "번째 쿼리 실행"
            strQuery = arrQuery(i)

            If strType = DB_TYPE_EXCEL And InStr(strQuery, "INSERT") > 0 Then
                strQuery = Replace(strQuery, "$]", "$1:" & lngEndRow + j & "]")
                j = j + 1
            End If

            Err.Clear
            .CommandText = strQuery
            .Execute
        Next i
    End With

    ExecuteQueryArr = True
    GoTo CleanUp

ErrHandler:
    intRetry = intRetry + 1

    If InStr(strErrLoc, "워크시트") > 0 Then
        ExecuteQueryArr = False
        GoTo ShowErr
    End If

    If intRetry <= intMaxRetry Then
        On Error Resume Next
        If Not objConn Is Nothing Then
            If objConn.State = 1 Then objConn.Close
        End If
        Set objCmd = Nothing
        Set objConn = Nothing
        On Error GoTo 0
        Application.Wait Now + TimeValue("00:00:01")
        GoTo RetryStart
    End If

ShowErr:
    ExecuteQueryArr = False
    MsgBox "오류 위치 : " & strErrLoc & vbCrLf & vbCrLf & _
           "재시도    : " & intRetry - 1 & "/" & intMaxRetry & vbCrLf & _
           "오류 번호 : " & Err.Number & vbCrLf & _
           "오류 내용 : " & Err.Description & vbCrLf & vbCrLf & _
           "쿼리      : " & strQuery, vbCritical, am_Core.AM_NAME

CleanUp:
    On Error Resume Next
    If Not objConn Is Nothing Then
        If objConn.State = 1 Then objConn.Close
    End If
    On Error GoTo 0
    Erase arrQuery
    Set objCmd = Nothing
    Set objConn = Nothing
    Set ws = Nothing
    Set rngInsert = Nothing

End Function

' 목적   : 조회 쿼리 결과를 셀 범위에 직접 출력
' 인수   : rngTarget    - 출력 시작 셀
'          arrQuery     - 실행할 쿼리 배열
'          strType      - DB 종류
'          strFile      - 파일 경로
'          strServer    - 서버 주소
'          strPort      - 포트 번호
'          strDB        - DB 명
'          strID        - 접속 ID
'          strPW        - 접속 PW
'          blnTranspose - True: 행열 전환 출력
'          intMoveCells - 쿼리 간격 이동 셀 수 (0: 자동)
'          blnHeader    - True: 헤더 출력 포함
' 예시   : SelectQuery(Sheet1.Range("A1"), Array("SELECT * FROM T1"), "엑세스", "C:\DB\main.accdb")
Public Sub SelectQuery(ByVal rngTarget As Range, _
                       ByVal arrQuery As Variant, _
                       ByVal strType As String, _
                       Optional ByVal strFile As String, _
                       Optional ByVal strServer As String, _
                       Optional ByVal strPort As String, _
                       Optional ByVal strDB As String, _
                       Optional ByVal strID As String, _
                       Optional ByVal strPW As String, _
                       Optional ByVal blnTranspose As Boolean = False, _
                       Optional ByVal intMoveCells As Integer = 0, _
                       Optional ByVal blnHeader As Boolean = False)

    Dim objConn     As Object
    Dim objRS       As Object
    Dim strConStr   As String
    Dim strQuery    As String
    Dim rngOut      As Range
    Dim intRetry    As Integer
    Dim intMaxRetry As Integer
    Dim i           As Long
    Dim j           As Long

    intMaxRetry = 3

RetryStart:
    On Error GoTo ErrHandler

    ' ── 1. DB 연결 ───────────────────────────────────────────
    Set objConn = CreateObject("ADODB.Connection")
    strConStr = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW, blnHeader)
    objConn.Open strConStr

    ' ── 2. 쿼리 실행 및 결과 출력 ───────────────────────────
    Set objRS = CreateObject("ADODB.Recordset")

    For i = LBound(arrQuery) To UBound(arrQuery)
        strQuery = CStr(arrQuery(i))
        objRS.Open strQuery, objConn

        If Not objRS.EOF Then
            If Not rngTarget Is Nothing Then

                Set rngOut = rngTarget

                If i > LBound(arrQuery) Then
                    If blnTranspose Then
                        If intMoveCells > 0 Then
                            Set rngOut = rngTarget.Offset(, (i - LBound(arrQuery)) * intMoveCells)
                        Else
                            With rngTarget.Worksheet
                                Set rngOut = .Cells(rngTarget.Row, .Columns.Count).End(xlToLeft)
                                If rngOut.Value2 <> "" Then Set rngOut = rngOut.Offset(, 1)
                            End With
                        End If
                    Else
                        If intMoveCells > 0 Then
                            Set rngOut = rngTarget.Offset((i - LBound(arrQuery)) * intMoveCells, 0)
                        Else
                            With rngTarget.Worksheet
                                Set rngOut = .Cells(.Rows.Count, rngTarget.Column).End(xlUp)
                                If rngOut.Value2 <> "" Then Set rngOut = rngOut.Offset(1, 0)
                            End With
                        End If
                    End If
                End If

                If Not rngOut.Worksheet Is ActiveSheet Then
                    rngOut.Worksheet.Activate
                End If

                If blnHeader Then
                    If blnTranspose Then
                        For j = 0 To objRS.Fields.Count - 1
                            rngOut.Offset(j, 0).Value2 = objRS.Fields(j).Name
                        Next j
                        Set rngOut = rngOut.Offset(0, 1)
                    Else
                        For j = 0 To objRS.Fields.Count - 1
                            rngOut.Offset(0, j).Value2 = objRS.Fields(j).Name
                        Next j
                        Set rngOut = rngOut.Offset(1, 0)
                    End If
                End If

                rngOut.CopyFromRecordset objRS

            End If
        End If

        objRS.Close
    Next i

    GoTo CleanUp

ErrHandler:
    intRetry = intRetry + 1

    If intRetry <= intMaxRetry Then
        On Error Resume Next
        If Not objRS Is Nothing Then
            If objRS.State = 1 Then objRS.Close
            Set objRS = Nothing
        End If
        If Not objConn Is Nothing Then
            If objConn.State = 1 Then objConn.Close
            Set objConn = Nothing
        End If
        On Error GoTo 0
        Application.Wait Now + TimeValue("00:00:01")
        GoTo RetryStart
    End If

    MsgBox "오류 번호 : " & Err.Number & vbCrLf & _
           "오류 내용 : " & Err.Description & vbCrLf & _
           "쿼리      : " & strQuery & vbCrLf & _
           "재시도    : " & intRetry - 1 & "/" & intMaxRetry, _
           vbCritical, am_Core.AM_NAME

CleanUp:
    On Error Resume Next
    If Not objRS Is Nothing Then
        If objRS.State = 1 Then objRS.Close
    End If
    If Not objConn Is Nothing Then
        If objConn.State = 1 Then objConn.Close
    End If
    On Error GoTo 0
    Set objRS = Nothing
    Set objConn = Nothing
    Set rngOut = Nothing

End Sub

' 목적   : 조회 쿼리 결과를 배열로 반환
' 인수   : arrQuery     - 실행할 쿼리 배열
'          strType      - DB 종류
'          strFile      - 파일 경로
'          strServer    - 서버 주소
'          strPort      - 포트 번호
'          strDB        - DB 명
'          strID        - 접속 ID
'          strPW        - 접속 PW
'          blnHeader    - True: 헤더 포함
'          blnTranspose - True: 행열 전환
' 반환   : Variant - 쿼리별 결과 배열 (배열의 배열)
' 예시   : arr = SelectQueryArr(Array("SELECT * FROM T1"), "엑세스", "C:\DB\main.accdb")
'          arr(0)(0, 0) → 첫번째 쿼리 첫번째 행 첫번째 열 값
Public Function SelectQueryArr(ByVal arrQuery As Variant, _
                               ByVal strType As String, _
                               Optional ByVal strFile As String, _
                               Optional ByVal strServer As String, _
                               Optional ByVal strPort As String, _
                               Optional ByVal strDB As String, _
                               Optional ByVal strID As String, _
                               Optional ByVal strPW As String, _
                               Optional ByVal blnHeader As Boolean = False, _
                               Optional ByVal blnTranspose As Boolean = False) As Variant

    Dim objConn      As Object
    Dim objRS        As Object
    Dim strConStr    As String
    Dim arrResult()  As Variant
    Dim arrTemp      As Variant
    Dim arrReTemp    As Variant
    Dim lngQuery     As Long
    Dim i            As Long
    Dim ri           As Long
    Dim ci           As Long
    Dim lngRow       As Long
    Dim lngCol       As Long
    Dim lngHdrOffset As Long
    Dim intRetry     As Integer
    Dim intMaxRetry  As Integer

    intMaxRetry = 3

RetryStart:
    Set objConn = CreateObject("ADODB.Connection")
    strConStr = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW, blnHeader)

    On Error GoTo ErrConn
    objConn.Open strConStr
    On Error GoTo 0

    ReDim arrResult(LBound(arrQuery) To UBound(arrQuery))

    For lngQuery = LBound(arrQuery) To UBound(arrQuery)
        Set objRS = CreateObject("ADODB.Recordset")

        On Error GoTo ErrQuery
        objRS.Open arrQuery(lngQuery), objConn
        On Error GoTo 0

        If Not objRS.EOF Then
            arrTemp = objRS.GetRows()

            For ri = UBound(arrTemp, 2) To 0 Step -1
                For ci = 0 To UBound(arrTemp, 1)
                    If Not IsNull(arrTemp(ci, ri)) Then GoTo FoundRow
                Next ci
            Next ri
FoundRow:
            lngRow = ri
            ReDim Preserve arrTemp(UBound(arrTemp, 1), lngRow)

            For ci = UBound(arrTemp, 1) To 0 Step -1
                For ri = 0 To UBound(arrTemp, 2)
                    If Not IsNull(arrTemp(ci, ri)) Then GoTo FoundCol
                Next ri
            Next ci
FoundCol:
            lngCol = ci
            lngHdrOffset = IIf(blnHeader, 1, 0)

            If blnTranspose Then
                ReDim arrReTemp(lngRow + lngHdrOffset, lngCol)
                If blnHeader Then
                    For ci = 0 To lngCol
                        arrReTemp(0, ci) = objRS.Fields(ci).Name
                    Next ci
                End If
                For ri = 0 To lngRow
                    For ci = 0 To lngCol
                        arrReTemp(ri + lngHdrOffset, ci) = arrTemp(ci, ri)
                    Next ci
                Next ri
            Else
                ReDim arrReTemp(lngCol, lngRow + lngHdrOffset)
                If blnHeader Then
                    For ci = 0 To lngCol
                        arrReTemp(ci, 0) = objRS.Fields(ci).Name
                    Next ci
                End If
                For ri = 0 To lngRow
                    For ci = 0 To lngCol
                        arrReTemp(ci, ri + lngHdrOffset) = arrTemp(ci, ri)
                    Next ci
                Next ri
            End If

            arrResult(lngQuery) = arrReTemp
        Else
            arrResult(lngQuery) = Array()
        End If

        objRS.Close
        Set objRS = Nothing
        GoTo NextQuery

ErrQuery:
        MsgBox "쿼리 오류 (인덱스: " & lngQuery & ")" & vbCrLf & _
               arrQuery(lngQuery), vbCritical, am_Core.AM_NAME
        arrResult(lngQuery) = Array()
        If Not objRS Is Nothing Then
            If objRS.State = 1 Then objRS.Close
            Set objRS = Nothing
        End If

NextQuery:
    Next lngQuery

    objConn.Close
    Set objConn = Nothing
    SelectQueryArr = arrResult
    Exit Function

ErrConn:
    intRetry = intRetry + 1

    If intRetry <= intMaxRetry Then
        On Error Resume Next
        If Not objConn Is Nothing Then
            If objConn.State = 1 Then objConn.Close
            Set objConn = Nothing
        End If
        On Error GoTo 0
        Application.Wait Now + TimeValue("00:00:01")
        GoTo RetryStart
    End If

    MsgBox "DB 연결 실패 (" & intRetry - 1 & "회 재시도)" & vbCrLf & _
           "오류 번호 : " & Err.Number & vbCrLf & _
           "오류 내용 : " & Err.Description, vbCritical, am_Core.AM_NAME
    SelectQueryArr = Array()

End Function

' ══════════════════════════════════════════════════════════
'  쿼리 생성
' ══════════════════════════════════════════════════════════

' 목적   : INSERT / UPDATE / DELETE 쿼리 문자열 생성
' 인수   : strQType  - 쿼리 종류 ("I" / "U" / "D")
'          strTable  - 대상 테이블명
'          arrData   - 필드/값 배열 (arrData(0,i)=필드, arrData(1,i)=값)
'          strWhere  - WHERE 조건문
'          blnParens - True: 필드명에 [] 괄호 추가
' 반환   : String - 완성된 쿼리 문자열
' 예시   : UpsertQuery("I", "T1", arrData)
'          UpsertQuery("U", "T1", arrData, "ID=1")
'          UpsertQuery("D", "T1", , "ID=1")
Public Function UpsertQuery(ByVal strQType As String, _
                            ByVal strTable As String, _
                            Optional ByVal arrData As Variant, _
                            Optional ByVal strWhere As String, _
                            Optional ByVal blnParens As Boolean = False) As String

    If strQType <> "D" And prv_IsArrayEmpty(arrData) Then Exit Function

    Dim strSQL        As String
    Dim strField      As String
    Dim strValue      As String
    Dim strFields     As String
    Dim strValues     As String
    Dim strFieldsVals As String
    Dim i             As Long

    Select Case UCase(strQType)
        Case "U"
            For i = LBound(arrData, 2) To UBound(arrData, 2)
                strField = IIf(blnParens, "[" & arrData(0, i) & "]", arrData(0, i))
                strValue = arrData(1, i)
                strFieldsVals = strFieldsVals & IIf(strFieldsVals = "", "", ", ") & _
                                strField & " = " & strValue
            Next i
            strSQL = "UPDATE " & strTable & _
                     " SET " & strFieldsVals & _
                     IIf(strWhere = "", "", " WHERE " & strWhere) & ";"

        Case "I"
            For i = LBound(arrData, 2) To UBound(arrData, 2)
                strField = IIf(blnParens, "[" & arrData(0, i) & "]", arrData(0, i))
                strValue = arrData(1, i)
                strFields = strFields & IIf(strFields = "", "", ", ") & strField
                strValues = strValues & IIf(strValues = "", "", ", ") & strValue
            Next i
            strSQL = "INSERT INTO " & strTable & _
                     " (" & strFields & ") VALUES (" & strValues & ");"

        Case "D"
            strSQL = "DELETE FROM " & strTable & _
                     IIf(strWhere = "", "", " WHERE " & strWhere) & ";"

    End Select

    UpsertQuery = strSQL

End Function

' 목적   : 범위 데이터를 DB 타입에 맞는 SQL VALUES 형식으로 변환
' 인수   : rngTarget - 변환할 셀 범위
'          arrTypes  - 각 열별 DB 타입명 배열
' 반환   : String - SQL VALUES 형식 문자열
' 예시   : ConvertRangeToSQL(Sheet1.Range("A1:C3"), Array("Integer","VarChar","Date"))
Public Function ConvertRangeToSQL(ByVal rngTarget As Range, _
                                  ByVal arrTypes As Variant) As String

    Dim strResult  As String
    Dim strRowData As String
    Dim strSQLVal  As String
    Dim vntCell    As Variant
    Dim lngRow     As Long
    Dim lngCol     As Long

    For lngRow = 1 To rngTarget.Rows.Count
        strRowData = ""
        For lngCol = 1 To rngTarget.Columns.Count
            vntCell = rngTarget.Cells(lngRow, lngCol).Value

            If lngCol <= UBound(arrTypes) - LBound(arrTypes) + 1 Then
                strSQLVal = FormatValueForSQLByDBType(vntCell, _
                            arrTypes(LBound(arrTypes) + lngCol - 1))
            Else
                strSQLVal = "'" & Replace(CStr(vntCell), "'", "''") & "'"
            End If

            strRowData = strRowData & IIf(strRowData = "", "", ", ") & strSQLVal
        Next lngCol

        strResult = strResult & IIf(strResult = "", "", "," & vbCrLf) & "(" & strRowData & ")"
    Next lngRow

    ConvertRangeToSQL = strResult

End Function

' ══════════════════════════════════════════════════════════
'  스키마 조회
' ══════════════════════════════════════════════════════════

' 목적   : DB / 엑셀 파일의 테이블(시트) 목록 반환
' 인수   : strType     - DB 종류 ("엑셀" / "엑세스")
'          strFilePath - 파일 경로
' 반환   : Variant - 테이블명 배열
' 예시   : GetDbTables("엑세스", "C:\DB\main.accdb")
Public Function GetDbTables(ByVal strType As String, _
                            ByVal strFilePath As String) As Variant

    Dim objConn     As Object
    Dim objRS       As Object
    Dim strConStr   As String
    Dim arrResult() As String
    Dim i           As Long

    Set objConn = CreateObject("ADODB.Connection")

    Select Case strType
        Case DB_TYPE_EXCEL
            strConStr = "Provider=Microsoft.ACE.OLEDB." & DB_EXCEL_VER & ";" & _
                        "Data Source=" & strFilePath & ";" & _
                        "Extended Properties=Excel " & DB_EXCEL_VER & ";"
        Case DB_TYPE_ACCESS
            strConStr = "Provider=Microsoft.ACE.OLEDB." & DB_ACCESS_VER & ";" & _
                        "Data Source=" & strFilePath & ";"
        Case Else
            MsgBox "'엑셀' 또는 '엑세스' 만 지원합니다.", vbCritical, am_Core.AM_NAME
            Exit Function
    End Select

    objConn.Open strConStr
    Set objRS = objConn.OpenSchema(20)

    Do While Not objRS.EOF
        Select Case strType
            Case DB_TYPE_EXCEL
                If InStr(objRS.Fields("TABLE_NAME").Value, "$") > 0 And _
                   InStr(objRS.Fields("TABLE_NAME").Value, "'") = 0 Then
                    ReDim Preserve arrResult(i)
                    arrResult(i) = Replace(objRS.Fields("TABLE_NAME").Value, "$", "")
                    i = i + 1
                End If
            Case DB_TYPE_ACCESS
                If objRS.Fields("TABLE_TYPE").Value = "TABLE" And _
                   Left(objRS.Fields("TABLE_NAME").Value, 4) <> "MSys" Then
                    ReDim Preserve arrResult(i)
                    arrResult(i) = objRS.Fields("TABLE_NAME").Value
                    i = i + 1
                End If
        End Select
        objRS.MoveNext
    Loop

    objRS.Close
    objConn.Close
    Set objRS = Nothing
    Set objConn = Nothing

    GetDbTables = arrResult

End Function

' 목적   : 테이블 필드 정보 반환
'          반환 배열 구조: (행, 열)
'          열 0=순번 / 1=필드명 / 2=타입명 / 3=길이 / 4=NULL여부 / 5=데이터수
' 인수   : strType   - DB 종류
'          strTable  - 테이블명
'          strFile   - 파일 경로
'          strServer - 서버 주소
'          strPort   - 포트 번호
'          strDB     - DB 명
'          strID     - 접속 ID
'          strPW     - 접속 PW
' 반환   : Variant - 필드 정보 2차원 배열
' 예시   : arr = GetFieldInfo("엑세스", "T1", "C:\DB\main.accdb")
'          arr(0, 1) → 첫번째 필드명
Public Function GetFieldInfo(ByVal strType As String, _
                             ByVal strTable As String, _
                             Optional ByVal strFile As String, _
                             Optional ByVal strServer As String, _
                             Optional ByVal strPort As String, _
                             Optional ByVal strDB As String, _
                             Optional ByVal strID As String, _
                             Optional ByVal strPW As String) As Variant

    Dim objConn       As Object
    Dim objRS         As Object
    Dim objRSCount    As Object
    Dim objFld        As Object
    Dim strConStr     As String
    Dim strQuery      As String
    Dim strCntQuery   As String
    Dim arrResult()   As Variant
    Dim arrFldNames() As String
    Dim dicCount      As Object
    Dim lngFldCnt     As Long
    Dim i             As Long
    Dim j             As Long
    Dim k             As Long
    Dim l             As Long
    Dim vntTemp       As Variant

    Set objConn = CreateObject("ADODB.Connection")
    Set dicCount = CreateObject("Scripting.Dictionary")

    strConStr = conStr(strType, strFile, strServer, strPort, strDB, strID, strPW)
    objConn.Open strConStr

    Select Case strType

        Case DB_TYPE_SERVER
            strQuery = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE " & _
                       "FROM information_schema.COLUMNS " & _
                       "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = '" & strTable & "' " & _
                       "ORDER BY ORDINAL_POSITION;"

            Set objRS = objConn.Execute(strQuery)

            Do While Not objRS.EOF
                ReDim Preserve arrFldNames(lngFldCnt)
                arrFldNames(lngFldCnt) = objRS.Fields("COLUMN_NAME").Value
                lngFldCnt = lngFldCnt + 1
                objRS.MoveNext
            Loop
            objRS.Close

            If lngFldCnt > 0 Then
                strCntQuery = "SELECT "
                For i = 0 To lngFldCnt - 1
                    strCntQuery = strCntQuery & IIf(i > 0, ", ", "") & _
                                  "COUNT(`" & arrFldNames(i) & "`) AS cnt" & i
                Next i
                strCntQuery = strCntQuery & " FROM `" & strTable & "`"

                On Error Resume Next
                Set objRSCount = objConn.Execute(strCntQuery)
                If Err.Number = 0 And Not objRSCount.EOF Then
                    For i = 0 To lngFldCnt - 1
                        dicCount.Add arrFldNames(i), objRSCount.Fields("cnt" & i).Value
                    Next i
                End If
                If Not objRSCount Is Nothing Then objRSCount.Close
                On Error GoTo 0
            End If

            Set objRS = objConn.Execute(strQuery)
            i = 0
            Do While Not objRS.EOF
                ReDim Preserve arrResult(5, i)
                arrResult(0, i) = i + 1
                arrResult(1, i) = objRS.Fields("COLUMN_NAME").Value
                arrResult(2, i) = GetDataTypeName(objRS.Fields("DATA_TYPE").Value)
                arrResult(3, i) = IIf(IsNull(objRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value), "", _
                                      objRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value)
                arrResult(4, i) = objRS.Fields("IS_NULLABLE").Value
                arrResult(5, i) = IIf(dicCount.Exists(objRS.Fields("COLUMN_NAME").Value), _
                                      dicCount(objRS.Fields("COLUMN_NAME").Value), 0)
                i = i + 1
                objRS.MoveNext
            Loop
            objRS.Close

        Case DB_TYPE_EXCEL
            If Right(strTable, 1) <> "$" Then strTable = strTable & "$"
            strQuery = "SELECT * FROM [" & strTable & "]"
            Set objRS = objConn.Execute(strQuery)

            lngFldCnt = 0
            ReDim arrFldNames(objRS.Fields.Count - 1)
            For Each objFld In objRS.Fields
                If Not prv_IsAutoGeneratedFieldName(objFld.Name) And _
                   Trim(objFld.Name) <> "" Then
                    arrFldNames(lngFldCnt) = objFld.Name
                    lngFldCnt = lngFldCnt + 1
                End If
            Next objFld

            If lngFldCnt > 0 Then
                ReDim Preserve arrFldNames(lngFldCnt - 1)
                strCntQuery = "SELECT "
                For i = 0 To lngFldCnt - 1
                    strCntQuery = strCntQuery & IIf(i > 0, ", ", "") & _
                                  "COUNT([" & arrFldNames(i) & "]) AS cnt" & i
                Next i
                strCntQuery = strCntQuery & " FROM [" & strTable & "]"

                On Error Resume Next
                Set objRSCount = objConn.Execute(strCntQuery)
                If Err.Number = 0 And Not objRSCount.EOF Then
                    For i = 0 To lngFldCnt - 1
                        dicCount.Add arrFldNames(i), objRSCount.Fields("cnt" & i).Value
                    Next i
                End If
                If Not objRSCount Is Nothing Then objRSCount.Close
                On Error GoTo 0
            End If

            i = 0
            For Each objFld In objRS.Fields
                If Not prv_IsAutoGeneratedFieldName(objFld.Name) And _
                   Trim(objFld.Name) <> "" Then
                    ReDim Preserve arrResult(5, i)
                    arrResult(0, i) = ""
                    arrResult(1, i) = objFld.Name
                    arrResult(2, i) = GetDataTypeName(objFld.Type)
                    arrResult(3, i) = IIf(objFld.DefinedSize > 0, objFld.DefinedSize, "")
                    arrResult(4, i) = ""
                    arrResult(5, i) = IIf(dicCount.Exists(objFld.Name), dicCount(objFld.Name), 0)
                    i = i + 1
                End If
            Next objFld
            objRS.Close

        Case DB_TYPE_ACCESS
            Set objRS = objConn.OpenSchema(4, Array(Empty, Empty, strTable))

            lngFldCnt = 0
            Do While Not objRS.EOF
                ReDim Preserve arrFldNames(lngFldCnt)
                arrFldNames(lngFldCnt) = objRS.Fields("COLUMN_NAME").Value
                lngFldCnt = lngFldCnt + 1
                objRS.MoveNext
            Loop

            If lngFldCnt > 0 Then
                strCntQuery = "SELECT "
                For i = 0 To lngFldCnt - 1
                    strCntQuery = strCntQuery & IIf(i > 0, ", ", "") & _
                                  "COUNT([" & arrFldNames(i) & "]) AS cnt" & i
                Next i
                strCntQuery = strCntQuery & " FROM [" & strTable & "]"

                On Error Resume Next
                Set objRSCount = objConn.Execute(strCntQuery)
                If Err.Number = 0 And Not objRSCount.EOF Then
                    For i = 0 To lngFldCnt - 1
                        dicCount.Add arrFldNames(i), objRSCount.Fields("cnt" & i).Value
                    Next i
                End If
                If Not objRSCount Is Nothing Then objRSCount.Close
                On Error GoTo 0
            End If

            objRS.Close
            Set objRS = objConn.OpenSchema(4, Array(Empty, Empty, strTable))

            Dim arrTmp() As Variant
            i = 0
            Do While Not objRS.EOF
                ReDim Preserve arrTmp(6, i)
                arrTmp(0, i) = objRS.Fields("ORDINAL_POSITION").Value
                arrTmp(1, i) = objRS.Fields("COLUMN_NAME").Value
                arrTmp(2, i) = GetDataTypeName(objRS.Fields("DATA_TYPE").Value)
                arrTmp(3, i) = IIf(IsNull(objRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value), "", _
                                   objRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value)
                arrTmp(4, i) = IIf(IsNull(objRS.Fields("IS_NULLABLE").Value), "", _
                                   objRS.Fields("IS_NULLABLE").Value)
                arrTmp(5, i) = objRS.Fields("ORDINAL_POSITION").Value
                arrTmp(6, i) = IIf(dicCount.Exists(objRS.Fields("COLUMN_NAME").Value), _
                                   dicCount(objRS.Fields("COLUMN_NAME").Value), 0)
                i = i + 1
                objRS.MoveNext
            Loop
            objRS.Close

            If i > 0 Then
                For j = 0 To i - 2
                    For k = j + 1 To i - 1
                        If arrTmp(5, j) > arrTmp(5, k) Then
                            For l = 0 To 6
                                vntTemp = arrTmp(l, j)
                                arrTmp(l, j) = arrTmp(l, k)
                                arrTmp(l, k) = vntTemp
                            Next l
                        End If
                    Next k
                Next j

                ReDim arrResult(5, i - 1)
                For j = 0 To i - 1
                    arrResult(0, j) = arrTmp(0, j)
                    arrResult(1, j) = arrTmp(1, j)
                    arrResult(2, j) = arrTmp(2, j)
                    arrResult(3, j) = arrTmp(3, j)
                    arrResult(4, j) = arrTmp(4, j)
                    arrResult(5, j) = arrTmp(6, j)
                Next j
            End If

    End Select

    objConn.Close
    Set objRS = Nothing
    Set objRSCount = Nothing
    Set objConn = Nothing
    Set dicCount = Nothing

    Dim arrFinal() As Variant
    Dim lngRows    As Long
    Dim lngCols    As Long
    Dim m          As Long
    Dim n          As Long

    On Error Resume Next
    lngRows = UBound(arrResult, 2) - LBound(arrResult, 2) + 1
    lngCols = UBound(arrResult, 1) - LBound(arrResult, 1) + 1
    On Error GoTo 0

    If lngRows > 0 And lngCols > 0 Then
        ReDim arrFinal(0 To lngRows - 1, 0 To lngCols - 1)
        For m = LBound(arrResult, 1) To UBound(arrResult, 1)
            For n = LBound(arrResult, 2) To UBound(arrResult, 2)
                arrFinal(n, m) = arrResult(m, n)
            Next n
        Next m
        GetFieldInfo = arrFinal
    Else
        GetFieldInfo = Array()
    End If

End Function

' ══════════════════════════════════════════════════════════
'  타입 처리
' ══════════════════════════════════════════════════════════

' 목적   : ADO 타입 코드를 타입명 문자열로 변환
' 인수   : intTypeCode - ADO 정수형 타입 코드
' 반환   : String - 타입명
' 예시   : GetDataTypeName(3) → "Integer"
Public Function GetDataTypeName(ByVal intTypeCode As Integer) As String
    Select Case intTypeCode
        Case 2:   GetDataTypeName = "SmallInt"
        Case 3:   GetDataTypeName = "Integer"
        Case 4:   GetDataTypeName = "Single"
        Case 5:   GetDataTypeName = "Double"
        Case 6:   GetDataTypeName = "Currency"
        Case 7:   GetDataTypeName = "Date"
        Case 11:  GetDataTypeName = "Boolean"
        Case 17:  GetDataTypeName = "TinyInt"
        Case 20:  GetDataTypeName = "BigInt"
        Case 72:  GetDataTypeName = "GUID"
        Case 128: GetDataTypeName = "Binary"
        Case 129: GetDataTypeName = "Char"
        Case 130: GetDataTypeName = "VarChar"
        Case 131: GetDataTypeName = "Numeric"
        Case 200: GetDataTypeName = "VarChar"
        Case 201: GetDataTypeName = "LongText"
        Case 202: GetDataTypeName = "VarWChar"
        Case 203: GetDataTypeName = "LongVarWChar"
        Case 204: GetDataTypeName = "VarBinary"
        Case 205: GetDataTypeName = "LongVarBinary"
        Case Else: GetDataTypeName = "Unknown(" & intTypeCode & ")"
    End Select
End Function

' 목적   : 타입명 문자열을 ADO 타입 코드로 변환
' 인수   : strTypeName - 타입명 문자열
' 반환   : Integer - ADO 정수형 타입 코드, 미지원 시 -1
' 예시   : GetDataTypeCode("Integer") → 3
Public Function GetDataTypeCode(ByVal strTypeName As String) As Integer
    Select Case UCase(Trim(strTypeName))
        Case "SMALLINT":               GetDataTypeCode = 2
        Case "INTEGER", "LONG", "INT": GetDataTypeCode = 3
        Case "SINGLE":                 GetDataTypeCode = 4
        Case "DOUBLE":                 GetDataTypeCode = 5
        Case "CURRENCY":               GetDataTypeCode = 6
        Case "DATE", "DATETIME":       GetDataTypeCode = 7
        Case "BOOLEAN", "YESNO":       GetDataTypeCode = 11
        Case "TINYINT":                GetDataTypeCode = 17
        Case "BIGINT":                 GetDataTypeCode = 20
        Case "GUID":                   GetDataTypeCode = 72
        Case "BINARY":                 GetDataTypeCode = 128
        Case "CHAR":                   GetDataTypeCode = 129
        Case "VARCHAR":                GetDataTypeCode = 130
        Case "NUMERIC":                GetDataTypeCode = 131
        Case "LONGTEXT", "MEMO":       GetDataTypeCode = 201
        Case "VARWCHAR", "TEXT":       GetDataTypeCode = 202
        Case "LONGVARWCHAR":           GetDataTypeCode = 203
        Case "VARBINARY":              GetDataTypeCode = 204
        Case "LONGVARBINARY":          GetDataTypeCode = 205
        Case Else:                     GetDataTypeCode = -1
    End Select
End Function

' 목적   : 지정 DB 타입에 맞는 SQL 값 문자열로 변환
' 인수   : vntValue    - 변환할 값
'          strTypeName - DB 타입명
' 반환   : String - SQL 값 문자열 (예: 'text', 123, NULL)
' 예시   : FormatValueForSQLByDBType("홍길동", "VarChar") → "'홍길동'"
Public Function FormatValueForSQLByDBType(ByVal vntValue As Variant, _
                                          ByVal strTypeName As String) As String

    If IsEmpty(vntValue) Or IsNull(vntValue) Or vntValue = "" Then
        FormatValueForSQLByDBType = "NULL"
        Exit Function
    End If

    Dim strType As String
    strType = UCase(Trim(strTypeName))

    Select Case strType
        Case "SMALLINT", "INTEGER", "TINYINT", "BIGINT"
            FormatValueForSQLByDBType = IIf(IsNumeric(vntValue), CStr(CLng(vntValue)), "NULL")

        Case "SINGLE", "DOUBLE"
            FormatValueForSQLByDBType = IIf(IsNumeric(vntValue), CStr(CDbl(vntValue)), "NULL")

        Case "CURRENCY"
            FormatValueForSQLByDBType = IIf(IsNumeric(vntValue), _
                                            CStr(Round(CDbl(vntValue), 4)), "NULL")

        Case "NUMERIC"
            FormatValueForSQLByDBType = IIf(IsNumeric(vntValue), CStr(CDbl(vntValue)), "NULL")

        Case "DATE"
            If IsDate(vntValue) Then
                FormatValueForSQLByDBType = "'" & Format(CDate(vntValue), "yyyy-mm-dd hh:nn:ss") & "'"
            ElseIf IsNumeric(vntValue) Then
                On Error Resume Next
                Dim dtVal As Date
                dtVal = CDate(CDbl(vntValue))
                FormatValueForSQLByDBType = IIf(Err.Number = 0, _
                    "'" & Format(dtVal, "yyyy-mm-dd hh:nn:ss") & "'", "NULL")
                On Error GoTo 0
            Else
                FormatValueForSQLByDBType = "NULL"
            End If

        Case "BOOLEAN"
            If VarType(vntValue) = vbBoolean Then
                FormatValueForSQLByDBType = IIf(CBool(vntValue), "True", "False")
            ElseIf IsNumeric(vntValue) Then
                FormatValueForSQLByDBType = IIf(CDbl(vntValue) <> 0, "True", "False")
            ElseIf UCase(CStr(vntValue)) = "TRUE" Or UCase(CStr(vntValue)) = "FALSE" Then
                FormatValueForSQLByDBType = UCase(CStr(vntValue))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If

        Case "GUID"
            Dim strGuid As String
            strGuid = CStr(vntValue)
            If Len(strGuid) >= 32 Then
                If Left(strGuid, 1) <> "{" Then strGuid = "{" & strGuid
                If Right(strGuid, 1) <> "}" Then strGuid = strGuid & "}"
                FormatValueForSQLByDBType = "'" & strGuid & "'"
            Else
                FormatValueForSQLByDBType = "NULL"
            End If

        Case "BINARY", "VARBINARY", "LONGVARBINARY"
            FormatValueForSQLByDBType = "NULL"

        Case Else
            FormatValueForSQLByDBType = "'" & Replace(CStr(vntValue), "'", "''") & "'"

    End Select

End Function

' 목적   : 지정 DB 타입과 일치하는지 검사
' 인수   : vntValue    - 검사할 값
'          strTypeName - DB 타입명
' 반환   : Boolean - True: 유효 / False: 유효하지 않음
' 예시   : ValidateValueForDBType(123, "Integer") → True
Public Function ValidateValueForDBType(ByVal vntValue As Variant, _
                                       ByVal strTypeName As String) As Boolean

    If IsEmpty(vntValue) Or IsNull(vntValue) Or vntValue = "" Then
        ValidateValueForDBType = True
        Exit Function
    End If

    Dim strType As String
    strType = UCase(Trim(strTypeName))

    Select Case strType
        Case "SMALLINT", "INTEGER", "TINYINT", "BIGINT"
            ValidateValueForDBType = IsNumeric(vntValue)

        Case "SINGLE", "DOUBLE", "CURRENCY", "NUMERIC"
            ValidateValueForDBType = IsNumeric(vntValue)

        Case "DATE"
            If IsDate(vntValue) Then
                ValidateValueForDBType = True
            ElseIf IsNumeric(vntValue) Then
                Dim dblVal As Double
                dblVal = CDbl(vntValue)
                ValidateValueForDBType = (dblVal >= 1 And dblVal <= 2958465)
            End If

        Case "BOOLEAN"
            ValidateValueForDBType = (VarType(vntValue) = vbBoolean) Or _
                                     IsNumeric(vntValue) Or _
                                     (UCase(CStr(vntValue)) = "TRUE") Or _
                                     (UCase(CStr(vntValue)) = "FALSE")

        Case "GUID"
            ValidateValueForDBType = (Len(CStr(vntValue)) >= 32)

        Case Else
            ValidateValueForDBType = True

    End Select

End Function

' ══════════════════════════════════════════════════════════
'  보조 프로시저 (내부 전용 - Private)
' ══════════════════════════════════════════════════════════

' 목적   : 자동 생성 필드명 여부 확인 (F1, F12, F123 형식)
Private Function prv_IsAutoGeneratedFieldName(ByVal strFieldName As String) As Boolean
    prv_IsAutoGeneratedFieldName = (strFieldName Like "F[0-9]" Or _
                                    strFieldName Like "F[0-9][0-9]" Or _
                                    strFieldName Like "F[0-9][0-9][0-9]")
End Function

' 목적   : 배열이 비어있는지 확인
Private Function prv_IsArrayEmpty(ByVal arr As Variant) As Boolean
    On Error Resume Next
    prv_IsArrayEmpty = (UBound(arr) < LBound(arr))
    If Err.Number <> 0 Then prv_IsArrayEmpty = True
    On Error GoTo 0
End Function

' 목적   : 빈 셀 자동 ID 채움 (현재 최댓값 이후로 순서 부여)
' 인수   : rng - ID 를 채울 셀 범위
' 예시   : AutoIDs(Sheet1.Range("A2:A100"))
Public Sub AutoIDs(ByVal rng As Range)

    Dim arrIDs()  As Long
    Dim r         As Range
    Dim i         As Long
    Dim lngMaxID  As Long
    Dim lngNextID As Long

    lngMaxID = Application.WorksheetFunction.Max(rng)
    lngNextID = lngMaxID + 1

    For Each r In rng
        ReDim Preserve arrIDs(i)
        arrIDs(i) = IIf(r.Value = "", lngNextID, r.Value)
        If r.Value = "" Then lngNextID = lngNextID + 1
        i = i + 1
    Next r

    rng.Value = Application.Transpose(arrIDs)

End Sub
