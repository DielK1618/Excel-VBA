Attribute VB_Name = "tpl_MySQL_Sub"
Option Explicit
Public Const PublicODBC_Driver As String = "MySQL ODBC 8.2 Unicode Driver"
Public Const PublicExcel_Ver As String = "12.0"
Public Const PublicAccess_Ver As String = "12.0"
Private p_Type As String
Private p_Token As String
Private p_File As String
Private p_Server As String
Private p_Port As String
Private p_DB As String
Private p_ID As String
Private p_PW As String
Private p_Initialized As Boolean
'자동 초기화를 포함한 Property
Public Property Get cfg_Type() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_Type = p_Type
End Property
Public Property Get cfg_Token() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_Token = p_Token
End Property
Public Property Get cfg_File() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_File = p_File
End Property
Public Property Get cfg_Server() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_Server = p_Server
End Property
Public Property Get cfg_Port() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_Port = p_Port
End Property
Public Property Get cfg_DB() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_DB = p_DB
End Property
Public Property Get cfg_ID() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_ID = p_ID
End Property
Public Property Get cfg_PW() As String
    If Not p_Initialized Then Call GetCfgDbInfo
    cfg_PW = p_PW
End Property
Sub GetCfgDbInfo()

    If p_Initialized Then Exit Sub ' 중복 실행 방지
    
    Dim tbl As ListObject
    Set tbl = GetTwbRange("T_데이터베이스관리").ListObject
    
    Dim vntToken As Variant
    vntToken = TblFindVals_MC(tbl, "예약어", "예약어", "=", "CFG_DB")
    
    Dim arrResult()
    arrResult = GetDbInfo(CStr(vntToken(0)))
    
    p_Type = arrResult(0)
    p_Token = arrResult(2)
    p_File = ReplacePath(arrResult(3))
    p_DB = arrResult(4)
    p_Server = arrResult(5)
    p_Port = arrResult(6)
    p_ID = arrResult(7)
    p_PW = arrResult(8)
    
    p_Initialized = True

End Sub
Function GetDbInfo(ByVal strToken As String, Optional ByVal boolToken = True) As Variant
    '----------------------------------------------------------------------------
    ' 반환은 다음의 순서대로 반환이 됩니다.
    ' 0,유형, 1.경로, 2.데이터베이스, 3.Server, 4.Port, 5.Id, 6.Pw
    ' GetDbinfo 뒤에 번호를 붙여 사용할 수 있습니다.
    ' GetDbinfo(0)은 경로입니다.
    '----------------------------------------------------------------------------
    Dim arr(8)
    Dim tbl As ListObject
    Set tbl = GetTwbRange("T_데이터베이스관리").ListObject
    
    Dim f As Range
    
    If boolToken Then
        Set f = TblFindRng_MC(tbl, "데이터베이스", "예약어", "=", strToken)
    Else
        Set f = TblFindRng_MC(tbl, "데이터베이스", "별칭", "=", strToken)
    End If
    
    If Not f Is Nothing Then
        arr(0) = f.Offset(, intOffset(f, "유형"))
        arr(1) = f.Offset(, intOffset(f, "별칭"))
        arr(2) = f.Offset(, intOffset(f, "예약어"))
        arr(3) = ReplacePath(f.Offset(, intOffset(f, "파일")))
        arr(4) = f.Offset(, intOffset(f, "데이터베이스"))
        arr(5) = f.Offset(, intOffset(f, "Server"))
        arr(6) = f.Offset(, intOffset(f, "Port"))
        arr(7) = f.Offset(, intOffset(f, "Id"))
        arr(8) = f.Offset(, intOffset(f, "Pw"))
    End If

    GetDbInfo = arr
    
End Function
Sub asdfadsfasfd()
    Dim tbl As ListObject
    Set tbl = GetTwbRange("T_데이터베이스관리").ListObject
    
    Dim f As Range
        Set f = TblFindRng_MC(tbl, "데이터베이스", "예약어", "=", "TST_DB")
        Debug.Print f.Address
        
End Sub
Sub ReloadDbInfo() ' 필요 시 재로드
    p_Initialized = False
    Call GetCfgDbInfo
End Sub
Sub GetFieldMap(Optional boolParens As Boolean = False)
    
    Dim arr As Variant
    Dim vntItem As Variant
    Dim strField, strID, strKey As String
    Dim i As Long
    Dim dv As DBvar
    
    With dv
        .Token = "CFG_DB"
        .Table = "cfg_fieldmaping"
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Token)
        .Type = arrDB(0)
        .File = arrDB(3)
    
        arr = SelectQueryArr(Array("SELECT * FROM " & .Table), .Type, .File)
    End With
    
    ' Dictionary 초기화
    Set dicPub = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    For i = 0 To UBound(arr(0), 2)
        If boolParens Then
            strField = "[" & arr(0)(3, i) & "]"
        Else
            strField = arr(0)(3, i)
        End If
        strID = arr(0)(0, i)
        vntItem = Array(strField, strID)
        strKey = arr(0)(1, i) & arr(0)(2, i) & arr(0)(4, i)
        dicPub.Add strKey, vntItem ' 키와 값의 순서 변경
    Next
    On Error GoTo 0
    
End Sub
Sub GetFieldMapRev(Optional boolParens As Boolean = False)
    
    Dim arr As Variant
    Dim vntItem As Variant
    Dim strField, strID, strKey As String
    Dim i As Long
    Dim dv As DBvar
    
    With dv
        .Token = "CFG_DB"
        .Table = "cfg_fieldmaping"
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Token)
        .Type = arrDB(0)
        .File = arrDB(3)
    
        arr = SelectQueryArr(Array("SELECT * FROM " & .Table), .Type, .File)
    End With
    
    ' Dictionary 초기화
    Set dicPubRev = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    For i = 0 To UBound(arr(0), 2)
        
        If boolParens Then
            strField = "[" & arr(0)(4, i) & "]"
        Else
            strField = arr(0)(4, i)
        End If
        strID = arr(0)(0, i)
        vntItem = Array(strField, strID)
        strKey = arr(0)(1, i) & arr(0)(2, i) & arr(0)(3, i)
        dicPubRev.Add strKey, vntItem
        
    Next
    On Error GoTo 0
    
End Sub
Function ReplaceFields(ByVal arrFields As Variant, _
                                            Optional strToken As String, _
                                            Optional strTable, _
                                            Optional boolParens As Boolean = False, _
                                            Optional boolRev As Boolean = False, _
                                            Optional boolNull As Boolean = False) As Variant
    
    If IsArrayEmpty(arrFields) Then Exit Function
    If boolRev Then
        If dicPubRev Is Nothing Then Call GetFieldMapRev(boolParens)
    Else
        If dicPub Is Nothing Then Call GetFieldMap(boolParens)
    End If
    
    Dim arr()
    Dim strKey As String
    Dim i As Integer
    
    For i = 0 To UBound(arrFields)
        ReDim Preserve arr(i)
        strKey = strToken & strTable & arrFields(i)
        
        If boolRev Then
            If dicPubRev.Exists(strKey) Then
                arr(i) = dicPubRev(strKey)(0)
            Else
                arr(i) = IIf(boolNull, "", arrFields(i))
            End If
        Else
            If dicPub.Exists(strKey) Then
                arr(i) = dicPub(strKey)(0)
            Else
                arr(i) = IIf(boolNull, "", arrFields(i))
            End If
        End If
    Next
    
    ReplaceFields = arr
    
End Function
Function GetFields(Optional intFontColor As Long, _
                                   Optional intBackGroundColor As Long, _
                                   Optional tbl As ListObject) As Variant
    
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    
    Dim arr()
    Dim r As Range
    Dim strField, strFields As String
    Dim i As Long
    
    For Each r In tbl.HeaderRowRange

        If intFontColor <> 0 Then If IsNull(r.Font.Color) Or r.Font.Color <> intFontColor Then GoTo pass
        If intBackGroundColor <> 0 Then If r.Interior.Color <> intBackGroundColor Then GoTo pass

        ReDim Preserve arr(i)
        
        arr(i) = r
        i = i + 1
pass:
    Next
    
    GetFields = arr
End Function
''<< 메인 코드 >>
Function FormatValueForSQLByDBType(dataValue As Variant, strTypeName As String) As String
    ' 데이터베이스 필드 타입에 맞춰 값을 SQL 형식으로 변환하는 함수
    ' strTypeName: "SmallInt", "Integer", "VarChar" 등의 타입명 문자열
    
    Dim upperTypeName As String
    upperTypeName = UCase(Trim(strTypeName))
    
    ' 빈 값이나 NULL 처리
    If IsEmpty(dataValue) Or IsNull(dataValue) Or dataValue = "" Then
        FormatValueForSQLByDBType = "NULL"
        Exit Function
    End If
    
    Select Case upperTypeName
        Case "SMALLINT", "INTEGER", "TINYINT", "BIGINT"  ' 정수형
            ' 정수형 - 숫자만 허용
            If IsNumeric(dataValue) Then
                ' 소수점 제거하고 정수로 변환
                FormatValueForSQLByDBType = CStr(CLng(dataValue))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "SINGLE", "DOUBLE"  ' 실수형
            ' 실수형 - 숫자만 허용
            If IsNumeric(dataValue) Then
                FormatValueForSQLByDBType = CStr(CDbl(dataValue))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "CURRENCY"  ' 통화형
            ' 통화형 - 숫자만 허용, 소수점 4자리까지
            If IsNumeric(dataValue) Then
                FormatValueForSQLByDBType = CStr(Round(CDbl(dataValue), 4))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "NUMERIC"  ' Numeric (고정 정밀도)
            ' 숫자형 - 숫자만 허용
            If IsNumeric(dataValue) Then
                FormatValueForSQLByDBType = CStr(CDbl(dataValue))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "DATE"  ' 날짜/시간형
            ' 날짜/시간형 - 아포스트로피로 감싸서 반환
            ' 엑셀 날짜 시리얼 넘버(숫자)도 처리
            If IsDate(dataValue) Then
                FormatValueForSQLByDBType = "'" & Format(CDate(dataValue), "yyyy-mm-dd hh:nn:ss") & "'"
            ElseIf IsNumeric(dataValue) Then
                ' 숫자인 경우 엑셀 날짜 시리얼 넘버로 간주하고 날짜로 변환
                On Error Resume Next
                Dim dateVal As Date
                dateVal = CDate(CDbl(dataValue))
                If Err.Number = 0 Then
                    FormatValueForSQLByDBType = "'" & Format(dateVal, "yyyy-mm-dd hh:nn:ss") & "'"
                Else
                    FormatValueForSQLByDBType = "NULL"
                End If
                On Error GoTo 0
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "BOOLEAN"  ' 불린형
            ' 불린형 - True/False를 1/0 또는 True/False로 변환
            If VarType(dataValue) = vbBoolean Then
                ' Access의 경우 True/False 사용
                FormatValueForSQLByDBType = IIf(CBool(dataValue), "True", "False")
            ElseIf IsNumeric(dataValue) Then
                ' 숫자로 들어온 경우 0이 아니면 True
                FormatValueForSQLByDBType = IIf(CDbl(dataValue) <> 0, "True", "False")
            ElseIf UCase(CStr(dataValue)) = "TRUE" Or UCase(CStr(dataValue)) = "FALSE" Then
                FormatValueForSQLByDBType = UCase(CStr(dataValue))
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "GUID"  ' GUID형
            ' GUID형 - 중괄호로 감싸서 반환
            Dim guidStr As String
            guidStr = CStr(dataValue)
            ' GUID 형식 검증 (간단한 체크)
            If Len(guidStr) >= 32 Then
                If Left(guidStr, 1) <> "{" Then guidStr = "{" & guidStr
                If Right(guidStr, 1) <> "}" Then guidStr = guidStr & "}"
                FormatValueForSQLByDBType = "'" & guidStr & "'"
            Else
                FormatValueForSQLByDBType = "NULL"
            End If
            
        Case "BINARY", "VARBINARY", "LONGVARBINARY"  ' 이진 데이터
            ' 이진 데이터 - 16진수 문자열로 변환 필요 (실제 구현은 상황에 따라 다름)
            ' 일반적으로 이진 데이터는 별도 처리 필요
            FormatValueForSQLByDBType = "NULL"  ' 기본적으로 NULL 처리
            
        Case "CHAR", "VARCHAR", "SHORTTEXT"  ' 문자열형 (짧은 텍스트)
            ' 문자열형 - SQL Injection 방지를 위한 이스케이프 처리
            FormatValueForSQLByDBType = "'" & Replace(CStr(dataValue), "'", "''") & "'"
            
        Case "LONGTEXT", "LONGVARWCHAR", "MEMO"  ' 긴 텍스트
            ' 긴 텍스트형 - SQL Injection 방지를 위한 이스케이프 처리
            FormatValueForSQLByDBType = "'" & Replace(CStr(dataValue), "'", "''") & "'"
            
        Case "VARWCHAR", "NVARCHAR"  ' 유니코드 문자열
            ' 유니코드 문자열 - SQL Injection 방지를 위한 이스케이프 처리
            FormatValueForSQLByDBType = "'" & Replace(CStr(dataValue), "'", "''") & "'" 'FormatValueForSQLByDBType = "N'" & Replace(CStr(dataValue), "'", "''") & "'"
            
        Case Else  ' 알 수 없는 타입
            ' 알 수 없는 타입 - 기본적으로 텍스트로 처리
            FormatValueForSQLByDBType = "'" & Replace(CStr(dataValue), "'", "''") & "'"
    End Select
End Function

'<< 서브 코드 >>
Function ValidateValueForDBType(dataValue As Variant, strTypeName As String) As Boolean
    ' 데이터가 데이터베이스 타입과 일치하는지 검증하는 함수
    ' strTypeName: "SmallInt", "Integer", "VarChar" 등의 타입명 문자열
    
    Dim upperTypeName As String
    upperTypeName = UCase(Trim(strTypeName))
    
    ' 빈 값이나 NULL은 항상 유효
    If IsEmpty(dataValue) Or IsNull(dataValue) Or dataValue = "" Then
        ValidateValueForDBType = True
        Exit Function
    End If
    
    Select Case upperTypeName
        Case "SMALLINT", "INTEGER", "TINYINT", "BIGINT"  ' 정수형
            ValidateValueForDBType = IsNumeric(dataValue)
            
        Case "SINGLE", "DOUBLE", "CURRENCY", "NUMERIC"  ' 실수형, 통화형
            ValidateValueForDBType = IsNumeric(dataValue)
            
        Case "DATE"  ' 날짜형
            ' IsDate() 체크 또는 숫자(엑셀 시리얼 넘버)도 날짜로 인정
            If IsDate(dataValue) Then
                ValidateValueForDBType = True
            ElseIf IsNumeric(dataValue) Then
                ' 숫자인 경우 엑셀 날짜 시리얼 넘버로 간주 (유효 범위 체크)
                Dim numVal As Double
                numVal = CDbl(dataValue)
                ' 엑셀 날짜 범위: 1900-01-01(1) ~ 9999-12-31(약 2958465)
                ValidateValueForDBType = (numVal >= 1 And numVal <= 2958465)
            Else
                ValidateValueForDBType = False
            End If
            
        Case "BOOLEAN"  ' 불린형
            If VarType(dataValue) = vbBoolean Then
                ValidateValueForDBType = True
            ElseIf IsNumeric(dataValue) Then
                ValidateValueForDBType = True
            ElseIf UCase(CStr(dataValue)) = "TRUE" Or UCase(CStr(dataValue)) = "FALSE" Then
                ValidateValueForDBType = True
            Else
                ValidateValueForDBType = False
            End If
            
        Case "GUID"  ' GUID
            ' GUID 형식 간단 검증 (최소 32자 이상)
            ValidateValueForDBType = (Len(CStr(dataValue)) >= 32)
            
        Case "BINARY", "VARBINARY", "LONGVARBINARY"  ' 이진 데이터
            ' 이진 데이터는 별도 검증 필요
            ValidateValueForDBType = True
            
        Case "CHAR", "VARCHAR", "SHORTTEXT", "LONGTEXT", "LONGVARWCHAR", "MEMO", "VARWCHAR", "NVARCHAR"  ' 문자열형
            ' 문자열은 모든 값 허용
            ValidateValueForDBType = True
            
        Case Else
            ' 기타 타입은 일단 허용
            ValidateValueForDBType = True
    End Select
End Function
Function ConvertRangeToSQLWithDBTypes(TargetRange As Range, fieldTypes() As String) As String
    ' 범위의 모든 셀을 데이터베이스 타입에 맞춰 SQL 형식으로 변환하는 함수
    ' fieldTypes: 각 컬럼의 데이터베이스 타입명 배열 (예: "Integer", "VarChar", "Date")
    
    Dim result As String
    Dim rowData As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim currentRow As Long
    Dim cellValue As Variant
    Dim sqlValue As String
    
    result = ""
    currentRow = TargetRange.Row
    
    For rowIndex = 1 To TargetRange.Rows.count
        rowData = ""
        
        For colIndex = 1 To TargetRange.Columns.count
            cellValue = TargetRange.Cells(rowIndex, colIndex).value
            
            ' fieldTypes 배열 인덱스 체크
            If colIndex <= UBound(fieldTypes) - LBound(fieldTypes) + 1 Then
                sqlValue = FormatValueForSQLByDBType(cellValue, fieldTypes(LBound(fieldTypes) + colIndex - 1))
            Else
                ' 타입 정보가 없으면 텍스트로 처리
                sqlValue = "'" & Replace(CStr(cellValue), "'", "''") & "'"
            End If
            
            If rowData = "" Then
                rowData = sqlValue
            Else
                rowData = rowData & ", " & sqlValue
            End If
        Next colIndex
        
        If result = "" Then
            result = "(" & rowData & ")"
        Else
            result = result & "," & vbCrLf & "(" & rowData & ")"
        End If
    Next rowIndex
    
    ConvertRangeToSQLWithDBTypes = result
End Function
Function GetValueValidationReport(dataValue As Variant, strTypeName As String) As String
    ' 값 검증 결과를 자세히 보여주는 함수
    
    Dim isValid As Boolean
    Dim sqlValue As String
    Dim result As String
    
    isValid = ValidateValueForDBType(dataValue, strTypeName)
    sqlValue = FormatValueForSQLByDBType(dataValue, strTypeName)
    
    result = "=== 값 검증 보고서 ===" & vbCrLf
    result = result & "데이터베이스 타입: " & strTypeName & vbCrLf
    result = result & "입력 값: " & CStr(dataValue) & vbCrLf
    result = result & "값 타입: " & TypeName(dataValue) & vbCrLf
    result = result & "타입 일치: " & IIf(isValid, "예", "아니오") & vbCrLf
    result = result & "SQL 변환값: " & sqlValue
    
    GetValueValidationReport = result
End Function
Function GetFieldAndType(ByVal arrFields As Variant, _
                                                ByVal strType As String, _
                                                ByVal strToken As String, _
                                                ByVal strTable As String, _
                                                Optional ByVal strFile As String, _
                                                Optional ByVal strServer As String = "", _
                                                Optional ByVal strPort As String = "", _
                                                Optional ByVal strDB As String = "", _
                                                Optional ByVal strID As String = "", _
                                                Optional ByVal strPW As String = "", _
                                                Optional ByVal boolParens As Boolean = False, _
                                                Optional ByVal boolRev As Boolean = False) As Variant

    If IsArrayEmpty(arrFields) Then Exit Function
                                        
    Dim arr As Variant
    arr = GetFieldInfo(strType, strTable, strFile, strServer, strPort, strDB, strID, strPW)
    
    If IsArrayEmpty(arr) Then
        MsgBox "데이터베이스 정보를 확인하세요!", vbCritical
        Exit Function
    End If
    
    Dim dicFields As Object
    Set dicFields = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim strKey, strKeyRev, strValue As String
    
    On Error Resume Next
    For i = 0 To UBound(arr)
        strKey = arr(i, 1)
        strValue = arr(i, 2)
        dicFields.Add strKey, strValue
    Next
    On Error GoTo 0
    
    Erase arr
    
    For i = 0 To UBound(arrFields)
        ReDim Preserve arr(2, i)
        If boolRev Then
            strKey = arrFields(i)
            strKeyRev = ReplaceFields(Array(arrFields(i)), strToken, strTable, , True)(0)
        Else
            strKey = ReplaceFields(Array(arrFields(i)), strToken, strTable)(0)
            strKeyRev = arrFields(i)
        End If
        
        If dicFields.Exists(strKey) Then
            arr(0, i) = IIf(boolParens, "[" & strKey & "]", strKey)
            arr(1, i) = strKeyRev
            arr(2, i) = dicFields(strKey)
        Else
            Debug.Print strKey & "필드 오류!"
            Exit Function
        End If
    Next
    
    GetFieldAndType = arr
    
End Function
Function UpsertQuery(ByVal strType As String, _
                                          ByVal strTable As String, _
                                          Optional ByVal arrData As Variant, _
                                          Optional ByVal strWhere As String, _
                                          Optional ByVal boolParens As Boolean = False) As Variant
    
    ' 입력값 검증
    If strType <> "D" And IsArrayEmpty(arrData) Then
        UpsertQuery = ""
        Exit Function
    End If
    
    Dim strSQL, _
            strField, _
            strValue, _
            strFields, _
            strValues, _
            strFieldsAndValues As String
    Dim i As Long
    
    Select Case UCase(strType)
    Case "U" '업데이트
        For i = 0 To UBound(arrData, 2)
            strField = IIf(boolParens, "[" & arrData(0, i) & "]", arrData(0, i))
            strValue = arrData(1, i)
            strFieldsAndValues = strFieldsAndValues & IIf(strFieldsAndValues = "", strField & " = " & strValue, ", " & strField & " = " & strValue)
        Next
        strSQL = "UPDATE " & strTable & " SET " & strFieldsAndValues & IIf(strWhere = "", "", " WHERE " & strWhere) & ";"
    Case "I" '인설트
        For i = 0 To UBound(arrData, 2)
            strField = IIf(boolParens, "[" & arrData(0, i) & "]", arrData(0, i))
            strValue = arrData(1, i)
            
            strFields = strFields & IIf(strFields = "", strField, ", " & strField)
            strValues = strValues & IIf(strValues = "", strValue, ", " & strValue)
        Next
        
        strSQL = "INSERT INTO " & strTable & " (" & strFields & ") VALUES (" & strValues & ");"
    Case "D" '삭제
        strSQL = "DELETE FROM " & strTable & IIf(strWhere = "", "", " WHERE " & strWhere) & ";"
    End Select
        
        UpsertQuery = strSQL
        
End Function
Sub asdfasdfadsfasfd()
    Debug.Print FormatValueForSQLByDBType(Now, "Date")
End Sub
Sub RecDBUpdateDateTime(ByVal strDB As String, Optional ByVal strDateTime As String)
    
    If strDateTime = "" Then strDateTime = FormatValueForSQLByDBType(Now, "Date")
    
    Dim dv As DBvar
    With dv
        .Token = "CFG_DB"
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Token)
        
        .Type = arrDB(0)
        .File = arrDB(3)
        .db = arrDB(4)
        .Server = arrDB(5)
        .Port = arrDB(6)
        .ID = arrDB(7)
        .PW = arrDB(8)
        
        .Table = "cfg_updatelog"
        .QPlus = "[contents] = '" & strDB & "'"
        .Query = "SELECT COUNT(*) FROM " & .Table & " WHERE " & .QPlus
        .arrQuery = Array(.Query)
        
        Dim arr As Variant
        arr = SelectQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
        
        Erase .arrQuery
        
        If IsArrayEmpty(arr) Then Exit Sub
        If arr(0)(0, 0) > 0 Then
            ReDim .arrData(1, 0)
            .arrData(0, 0) = "[datetime]"
            .arrData(1, 0) = strDateTime
            
            .QType = "U"
        Else
            ReDim .arrData(1, 1)
            .arrData(0, 0) = "[contents]"
            .arrData(0, 1) = "[datetime]"
            .arrData(1, 0) = "'" & strDB & "'"
            .arrData(1, 1) = strDateTime
            
            .QType = "I"
        End If
        
        .arrQuery = Array(UpsertQuery(.QType, .Table, .arrData, .QPlus))
        
        If ExecuteQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW) = False Then
            Debug.Print "[" & strDB & "] 테이블의 업데이트 로그 기록 실패!"
        End If
    End With
    
    
    
End Sub
Function GetFieldNameConnection(ByVal strToken As String, _
                                              ByVal strTable As String) As Variant
                                              
   Dim dv As DBvar
    With dv
        .Token = strToken
        .Table = strTable
        
        If .Table = "" Then Exit Function

        Dim arr As Variant
        arr = GetDbInfo(.Token)
        
        .Type = arr(0)
        .File = arr(3)
        .db = arr(4)
        .Server = arr(5)
        .Port = arr(6)
        .ID = arr(7)
        .PW = arr(8)
                
        Erase arr
        arr = GetFieldInfo(.Type, .Table, .File, .Server, .Port, .db, .ID, .PW)
        
    If IsArrayEmpty(arr) Then Exit Function
        
    Dim arrValues()
    Dim i As Long
    
    For i = 0 To UBound(arr)
        ReDim Preserve arrValues(4, i)
        arrValues(0, i) = arr(i, 1)
        arrValues(1, i) = ReplaceFields(Array(arr(i, 1)), .Token, .Table, , True, True)(0)
        arrValues(2, i) = arr(i, 2)
        arrValues(3, i) = arr(i, 5)
        arrValues(4, i) = arr(i, 0)
    Next
    End With
    
    GetFieldNameConnection = arrValues
    
End Function
