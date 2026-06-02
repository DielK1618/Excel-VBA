Attribute VB_Name = "tpl_Access"
Option Explicit
' ============================================
' Access 테이블 관리 함수 - 최종 버전
' 작성일: 2026-01-20
' ============================================
' ============================================
' 1. 테이블 생성/수정 (SQL 방식 - 권장)
' ============================================
' 사용 예시:
' arrFields = Array(Array(1, "ID", 3), Array(2, "Name", 130), Array(3, "Age", 3))
' Call CreateAccessTable("C:\DB\config.accdb", "Users", arrFields)
Function CreateAccessTable(strFile As String, strTable As String, arrFieldAndTypes As Variant) As Boolean
    ' Access 데이터베이스에 테이블을 생성하는 함수 (SQL 방식)
    ' 인수:
    '   strFile - Access 파일 경로 (예: "C:\Data\MyDB.accdb")
    '   strTable - 테이블 이름
    '   arrFieldAndTypes - 2차원 배열 Array(순번, 필드명, 타입번호)
    '                      예: Array(Array(1, "ID", 3), Array(2, "Name", 130))
    ' 타입 번호:
    '   3 = Long(정수), 130/202/200 = Text(텍스트), 7 = Date(날짜)
    '   5 = Double(실수), 11 = Boolean(예/아니오), 12 = Variant(메모)
    ' 반환값: Boolean (성공 True, 실패 False)
    ' 동작:
    '   - 기존 테이블이 있으면 삭제 후 재생성
    '   - ID 필드는 자동으로 AUTOINCREMENT PRIMARY KEY 설정
    
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim strSQL As String
    Dim i As Integer
    Dim strFieldName As String
    Dim intFieldType As Integer
    Dim strFieldDef As String
    Dim arrFields() As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile
    
    ' 기존 테이블 삭제 (있으면)
    On Error Resume Next
    conn.Execute "DROP TABLE [" & strTable & "]"
    On Error GoTo ErrorHandler
    
    ' 필드 정의 생성
    ReDim arrFields(UBound(arrFieldAndTypes) - LBound(arrFieldAndTypes))
    
    For i = LBound(arrFieldAndTypes) To UBound(arrFieldAndTypes)
        strFieldName = arrFieldAndTypes(i)(1)
        intFieldType = arrFieldAndTypes(i)(2)
        
        ' 타입에 따른 SQL 정의
        Select Case intFieldType
            Case 3      ' Long Integer
                If UCase(strFieldName) = "ID" Then
                    strFieldDef = "[" & strFieldName & "] AUTOINCREMENT PRIMARY KEY"
                Else
                    strFieldDef = "[" & strFieldName & "] LONG"
                End If
            Case 130, 202, 200  ' Text
                strFieldDef = "[" & strFieldName & "] TEXT(255)"
            Case 7      ' Date/Time
                strFieldDef = "[" & strFieldName & "] DATETIME"
            Case 5      ' Double
                strFieldDef = "[" & strFieldName & "] DOUBLE"
            Case 11     ' Boolean
                strFieldDef = "[" & strFieldName & "] BIT"
            Case 12     ' Memo
                strFieldDef = "[" & strFieldName & "] MEMO"
            Case Else
                strFieldDef = "[" & strFieldName & "] TEXT(255)"
        End Select
        
        arrFields(i) = strFieldDef
    Next i
    
    ' CREATE TABLE SQL 생성 및 실행
    strSQL = "CREATE TABLE [" & strTable & "] (" & Join(arrFields, ", ") & ")"
    conn.Execute strSQL
    
    Debug.Print "테이블 '" & strTable & "' 생성 완료"
    CreateAccessTable = True
    
    conn.Close
    Set conn = Nothing
    Exit Function
    
ErrorHandler:
    CreateAccessTable = False
    Debug.Print "오류 발생: " & Err.Description
    
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Function
Function DeleteAccessTable(ByVal strFile As String, ByVal strTable As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cn As Object
    
    ' Access DB 연결
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";"
    
    ' 테이블의 모든 데이터 삭제
    cn.Execute "DELETE FROM " & strTable
    
    ' 연결 종료
    cn.Close
    Set cn = Nothing
    
    DeleteAccessTable = True
    Exit Function
    
ErrorHandler:
    DeleteAccessTable = False
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
    End If
    Set cn = Nothing
End Function
' ============================================
' 2. 테이블 생성/수정 (ADOX 방식 - 고급)
' ============================================
' 기존 테이블을 유지하면서 필드만 수정/추가하고 싶을 때 사용
Function CreateAccessTableADOX(strFile As String, strTable As String, arrFieldAndTypes As Variant) As Boolean
    ' Access 데이터베이스 테이블을 생성/수정하는 함수 (ADOX 방식)
    ' 동작:
    '   - 테이블이 없으면 새로 생성
    '   - 테이블이 있으면 순번 기반으로 필드 수정/추가
    '   - 기존 데이터 유지
    
    On Error GoTo ErrorHandler

    Dim cat As Object, tbl As Object, col As Object
    Dim conn As Object, rs As Object, idx As Object
    Dim i As Integer
    Dim intOrderNum As Integer, strFieldName As String, intFieldType As Integer
    Dim bTableExists As Boolean, strIDFieldName As String
    Dim dictExistingFields As Object
    Dim intExistingOrder As Integer, strExistingName As String, intExistingType As Integer

    Set dictExistingFields = CreateObject("Scripting.Dictionary")
    strIDFieldName = ""

    ' ADOX 카탈로그 및 ADO 연결
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile

    ' 테이블 존재 여부 확인
    bTableExists = False
    For i = 0 To cat.Tables.count - 1
        If cat.Tables(i).Name = strTable Then
            bTableExists = True
            Set tbl = cat.Tables(strTable)
            Exit For
        End If
    Next i

    ' 테이블이 없으면 새로 생성
    If Not bTableExists Then
        Set tbl = CreateObject("ADOX.Table")
        tbl.Name = strTable

        For i = LBound(arrFieldAndTypes) To UBound(arrFieldAndTypes)
            intOrderNum = arrFieldAndTypes(i)(0)
            strFieldName = arrFieldAndTypes(i)(1)
            intFieldType = arrFieldAndTypes(i)(2)

            Set col = CreateObject("ADOX.Column")
            col.Name = strFieldName
            col.Type = intFieldType

            ' ID 필드 AutoIncrement 설정 (안전한 방식)
            If UCase(strFieldName) = "ID" Then
                strIDFieldName = strFieldName
                On Error Resume Next
                col.Properties("AutoIncrement") = True
                If Err.Number <> 0 Then
                    Err.Clear
                    col.Properties("Jet OLEDB:AutoIncrement") = True
                    If Err.Number <> 0 Then Err.Clear
                End If
                On Error GoTo ErrorHandler
                col.Attributes = 2  ' NOT NULL
            End If

            ' 텍스트 타입 크기 지정
            If intFieldType = 202 Or intFieldType = 200 Or intFieldType = 130 Then
                col.DefinedSize = 255
            End If

            tbl.Columns.Append col
            Set col = Nothing
        Next i

        ' PK 설정
        If strIDFieldName <> "" Then
            Set idx = CreateObject("ADOX.Index")
            idx.Name = "PrimaryKey"
            idx.PrimaryKey = True
            idx.Unique = True

            Set col = CreateObject("ADOX.Column")
            col.Name = strIDFieldName
            idx.Columns.Append col

            tbl.Indexes.Append idx
            Set col = Nothing
            Set idx = Nothing
        End If

        cat.Tables.Append tbl
        Debug.Print "새 테이블 '" & strTable & "' 생성 완료"

    ' 테이블이 있으면 필드 수정/추가
    Else
        Debug.Print "기존 테이블 '" & strTable & "' 수정"
        
        ' 기존 필드 정보 수집
        Set rs = conn.OpenSchema(4, Array(Empty, Empty, strTable))
        Do While Not rs.EOF
            intExistingOrder = rs.fields("ORDINAL_POSITION").value
            strExistingName = rs.fields("COLUMN_NAME").value
            intExistingType = rs.fields("DATA_TYPE").value
            dictExistingFields.Add intExistingOrder, Array(strExistingName, intExistingType)
            rs.MoveNext
        Loop
        rs.Close

        ' 필드 수정/추가
        For i = LBound(arrFieldAndTypes) To UBound(arrFieldAndTypes)
            intOrderNum = arrFieldAndTypes(i)(0)
            strFieldName = arrFieldAndTypes(i)(1)
            intFieldType = arrFieldAndTypes(i)(2)

            If UCase(strFieldName) = "ID" Then strIDFieldName = strFieldName

            If dictExistingFields.Exists(intOrderNum) Then
                strExistingName = dictExistingFields(intOrderNum)(0)
                intExistingType = dictExistingFields(intOrderNum)(1)

                ' 이름이나 타입이 다르면 수정
                If (strExistingName <> strFieldName) Or (intExistingType <> intFieldType) Then
                    On Error Resume Next
                    Set col = tbl.Columns(strExistingName)
                    
                    If Err.Number = 0 Then
                        If strExistingName <> strFieldName Then
                            col.Name = strFieldName
                            Debug.Print "  필드명 변경: " & strExistingName & " → " & strFieldName
                        End If
                        
                        If intExistingType <> intFieldType Then
                            col.Type = intFieldType
                            If intFieldType = 202 Or intFieldType = 200 Or intFieldType = 130 Then
                                col.DefinedSize = 255
                            End If
                            Debug.Print "  타입 변경: " & intExistingType & " → " & intFieldType
                        End If
                    End If
                    
                    Err.Clear
                    On Error GoTo ErrorHandler
                    Set col = Nothing
                End If

            ' 해당 순번의 필드가 없으면 새로 추가
            Else
                Set col = CreateObject("ADOX.Column")
                col.Name = strFieldName
                col.Type = intFieldType

                If UCase(strFieldName) = "ID" Then
                    On Error Resume Next
                    col.Properties("AutoIncrement") = True
                    If Err.Number <> 0 Then
                        Err.Clear
                        col.Properties("Jet OLEDB:AutoIncrement") = True
                        If Err.Number <> 0 Then Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                    col.Attributes = 2
                End If

                If intFieldType = 202 Or intFieldType = 200 Or intFieldType = 130 Then
                    col.DefinedSize = 255
                End If

                tbl.Columns.Append col
                Debug.Print "  새 필드 추가: " & strFieldName
                Set col = Nothing
            End If
        Next i

        ' PK 추가 (없는 경우만)
        If strIDFieldName <> "" Then
            Dim bPKExists As Boolean
            bPKExists = False
            
            On Error Resume Next
            For i = 0 To tbl.Indexes.count - 1
                If tbl.Indexes(i).PrimaryKey = True Then
                    bPKExists = True
                    Exit For
                End If
            Next i
            On Error GoTo ErrorHandler

            If Not bPKExists Then
                On Error Resume Next
                Set idx = CreateObject("ADOX.Index")
                idx.Name = "PrimaryKey"
                idx.PrimaryKey = True
                idx.Unique = True

                Set col = CreateObject("ADOX.Column")
                col.Name = strIDFieldName
                idx.Columns.Append col
                tbl.Indexes.Append idx

                If Err.Number = 0 Then Debug.Print "  PK 설정: " & strIDFieldName
                Err.Clear
                Set col = Nothing
                Set idx = Nothing
                On Error GoTo ErrorHandler
            End If
        End If
    End If

    CreateAccessTableADOX = True

    Set rs = Nothing
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set tbl = Nothing
    Set cat = Nothing
    Set dictExistingFields = Nothing
    Exit Function

ErrorHandler:
    CreateAccessTableADOX = False
    Debug.Print "오류 발생: " & Err.Description

    If Not rs Is Nothing Then Set rs = Nothing
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not col Is Nothing Then Set col = Nothing
    If Not idx Is Nothing Then Set idx = Nothing
    If Not tbl Is Nothing Then Set tbl = Nothing
    If Not cat Is Nothing Then Set cat = Nothing
    If Not dictExistingFields Is Nothing Then Set dictExistingFields = Nothing
End Function
' ============================================
' 3. 필드 삭제 (순번 기반)
' ============================================
Function DeleteAccessFields(strFile As String, strTable As String, arrPosition As Variant) As Boolean
    ' 테이블의 특정 순번 필드를 삭제하는 함수
    ' 인수:
    '   arrPosition - 삭제할 필드 순번 배열 (예: Array(2, 5, 7))
    
    On Error GoTo ErrorHandler
    
    Dim cat As Object, tbl As Object
    Dim conn As Object, rs As Object
    Dim i As Integer, intDelPosition As Integer
    Dim dictExistingFields As Object
    Dim intExistingOrder As Integer, strExistingName As String
    Dim strFieldToDelete As String
    Dim intDeletedCount As Integer
    
    Set dictExistingFields = CreateObject("Scripting.Dictionary")
    intDeletedCount = 0
    
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile
    
    ' 테이블 존재 확인
    Dim bTableExists As Boolean
    bTableExists = False
    For i = 0 To cat.Tables.count - 1
        If cat.Tables(i).Name = strTable Then
            bTableExists = True
            Set tbl = cat.Tables(strTable)
            Exit For
        End If
    Next i
    
    If Not bTableExists Then
        Debug.Print "오류: 테이블 '" & strTable & "'이(가) 존재하지 않습니다."
        DeleteAccessFields = False
        GoTo CleanUp
    End If
    
    ' 기존 필드 순번 정보 수집
    Set rs = conn.OpenSchema(4, Array(Empty, Empty, strTable))
    Do While Not rs.EOF
        intExistingOrder = rs.fields("ORDINAL_POSITION").value
        strExistingName = rs.fields("COLUMN_NAME").value
        dictExistingFields.Add intExistingOrder, strExistingName
        rs.MoveNext
    Loop
    rs.Close
    
    ' 필드 삭제
    For i = LBound(arrPosition) To UBound(arrPosition)
        intDelPosition = arrPosition(i)
        
        If dictExistingFields.Exists(intDelPosition) Then
            strFieldToDelete = dictExistingFields(intDelPosition)
            
            On Error Resume Next
            tbl.Columns.Delete strFieldToDelete
            
            If Err.Number = 0 Then
                Debug.Print "순번 " & intDelPosition & ": 필드 '" & strFieldToDelete & "' 삭제 완료"
                intDeletedCount = intDeletedCount + 1
            Else
                Debug.Print "순번 " & intDelPosition & ": 필드 '" & strFieldToDelete & "' 삭제 실패 - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Else
            Debug.Print "순번 " & intDelPosition & ": 해당 필드 없음"
        End If
    Next i
    
    Debug.Print "필드 삭제 완료 (총 " & intDeletedCount & "개)"
    DeleteAccessFields = True
    
CleanUp:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not tbl Is Nothing Then Set tbl = Nothing
    If Not cat Is Nothing Then Set cat = Nothing
    If Not dictExistingFields Is Nothing Then Set dictExistingFields = Nothing
    Exit Function
    
ErrorHandler:
    DeleteAccessFields = False
    Debug.Print "오류 발생: " & Err.Description
    Resume CleanUp
End Function
' ============================================
' 4. 사용 예시
' ============================================
Sub Example_CreateTable()
    Dim strFile As String
    strFile = "D:\WORKSPACE\Repository\VBA\00_db\config.accdb"
    
    Dim strTable As String
    strTable = "cfg_fieldmaping"
    
    Dim arrFields()
    arrFields = Array(Array(1, "ID", 3), _
    Array(2, "token", 130), _
        Array(3, "tablename", 130), _
        Array(4, "fieldname", 130), _
        Array(5, "rename", 130) _
    )
    
    ' SQL 방식 (권장) - 간단하고 안정적
    If CreateAccessTable(strFile, strTable, arrFields) Then
        Debug.Print "성공!"
    End If
End Sub
Sub Example_UpdateTable()
    ' ADOX 방식 - 기존 데이터 유지하면서 필드 수정/추가
    Dim strFile As String
    strFile = "D:\WORKSPACE\Repository\VBA\00_db\config.accdb"
    
    Dim arrFields()
    arrFields = Array( _
        Array(1, "ID", 3), _
        Array(2, "token", 130), _
        Array(3, "tablename", 130), _
        Array(4, "fieldname", 130), _
        Array(5, "rename", 130), _
        Array(6, "description", 130) _
    )
    
    If CreateAccessTableADOX(strFile, "cfg_fieldmaping", arrFields) Then
        Debug.Print "수정 완료!"
    End If
End Sub
Sub Example_DeleteFields()
    ' 특정 순번의 필드 삭제
    Dim strFile As String
    strFile = "D:\WORKSPACE\Repository\VBA\00_db\config.accdb"
    
    ' 2번, 5번 순번의 필드 삭제
    If DeleteAccessFields(strFile, "cfg_fieldmaping", Array(2, 5)) Then
        Debug.Print "삭제 완료!"
    End If
End Sub
Sub DiagnoseAccessConnection() '디버깅용
    
    Dim filePath As String
    filePath = "D:\WORKSPACE\Repository\VBA\00_db\db_edulms.accdb"
    
    Debug.Print "=========================================="
    Debug.Print "Access 데이터베이스 연결 진단"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' 1. 파일 기본 정보
    Debug.Print "■ 1. 파일 기본 정보"
    Debug.Print "   경로: " & filePath
    Debug.Print "   존재 여부: " & (Dir(filePath) <> "")
    
    If Dir(filePath) = "" Then
        Debug.Print "   >>> 파일이 존재하지 않습니다!"
        Exit Sub
    End If
    
    Debug.Print "   크기: " & Format(FileLen(filePath) / 1024 / 1024, "0.00") & " MB"
    Debug.Print "   수정 날짜: " & FileDateTime(filePath)
    
    ' 2. 파일 속성 확인
    Debug.Print ""
    Debug.Print "■ 2. 파일 속성"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        Dim dbFile As Object
        Set dbFile = fso.GetFile(filePath)
        
        Debug.Print "   읽기 전용: " & CBool(dbFile.Attributes And 1)
        Debug.Print "   숨김: " & CBool(dbFile.Attributes And 2)
        Debug.Print "   시스템: " & CBool(dbFile.Attributes And 4)
        Debug.Print "   보관: " & CBool(dbFile.Attributes And 32)
    End If
    
    ' 3. 파일 잠금 상태
    Debug.Print ""
    Debug.Print "■ 3. 파일 잠금 상태"
    On Error Resume Next
    Dim testFile As Integer
    testFile = FreeFile
    Open filePath For Binary Access Read Lock Read Write As #testFile
    
    If Err.Number <> 0 Then
        Debug.Print "   상태: 잠김 (다른 프로세스에서 사용 중)"
        Debug.Print "   오류: " & Err.Description
        Err.Clear
    Else
        Debug.Print "   상태: 사용 가능"
        Close #testFile
    End If
    On Error GoTo 0
    
    ' 4. Access 파일 형식 확인
    Debug.Print ""
    Debug.Print "■ 4. Access 파일 형식"
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False, 0)
    Dim header As String
    header = ts.Read(100)
    ts.Close
    
    If InStr(header, "Standard ACE DB") > 0 Then
        Debug.Print "   형식: Access 2007 이상 (.accdb)"
    ElseIf InStr(header, "Standard Jet DB") > 0 Then
        Debug.Print "   형식: Access 2003 이하 (.mdb)"
    Else
        Debug.Print "   형식: 알 수 없음 (손상되었을 가능성)"
    End If
    
    ' 5. Office/Excel 버전 확인
    Debug.Print ""
    Debug.Print "■ 5. Office 환경"
    Debug.Print "   Excel 버전: " & Application.Version
    
    #If Win64 Then
        Debug.Print "   Excel 비트: 64-bit"
    #Else
        Debug.Print "   Excel 비트: 32-bit"
    #End If
    
    Debug.Print "   운영체제: " & Application.OperatingSystem
    
    ' 6. ADODB Provider 테스트
    Debug.Print ""
    Debug.Print "■ 6. ADODB Provider 테스트"
    
    Dim providers As Variant
    providers = Array( _
        "Microsoft.ACE.OLEDB.16.0", _
        "Microsoft.ACE.OLEDB.15.0", _
        "Microsoft.ACE.OLEDB.12.0", _
        "Microsoft.Jet.OLEDB.4.0" _
    )
    
    Dim p As Variant
    Dim cn As Object
    Dim connStr As String
    Dim successProvider As String
    
    For Each p In providers
        On Error Resume Next
        Set cn = CreateObject("ADODB.Connection")
        connStr = "Provider=" & p & ";Data Source=" & filePath & ";"
        cn.Open connStr
        
        If Err.Number = 0 And cn.State = 1 Then
            Debug.Print "   ? " & p & " - 성공!"
            successProvider = p
            cn.Close
            Set cn = Nothing
            Exit For
        Else
            Debug.Print "   ? " & p
            Debug.Print "      오류 번호: " & Err.Number
            Debug.Print "      오류 설명: " & Err.Description
            Err.Clear
            If Not cn Is Nothing Then
                If cn.State = 1 Then cn.Close
                Set cn = Nothing
            End If
        End If
    Next p
    
    On Error GoTo 0
    
    ' 7. DAO 테스트
    Debug.Print ""
    Debug.Print "■ 7. DAO 테스트"
    
    On Error Resume Next
    Dim db As Object
    Set db = CreateObject("DAO.DBEngine.120").OpenDatabase(filePath)
    
    If Err.Number = 0 Then
        Debug.Print "   ? DAO 연결 성공!"
        db.Close
        Set db = Nothing
    Else
        Debug.Print "   ? DAO 연결 실패"
        Debug.Print "      오류 번호: " & Err.Number
        Debug.Print "      오류 설명: " & Err.Description
        Err.Clear
    End If
    
    ' DAO 참조 확인
    Dim hasDAORef As Boolean
    hasDAORef = False
    
    Dim ref As Object
    For Each ref In ThisWorkbook.VBProject.References
        If InStr(LCase(ref.Description), "dao") > 0 Then
            hasDAORef = True
            Debug.Print "   DAO 참조: " & ref.Description & " (v" & ref.Major & "." & ref.Minor & ")"
        End If
    Next
    
    If Not hasDAORef Then
        Debug.Print "   DAO 참조: 없음 (도구 > 참조에서 추가 필요)"
    End If
    
    On Error GoTo 0
    
    ' 8. 설치된 OLEDB Provider 목록
    Debug.Print ""
    Debug.Print "■ 8. 시스템에 설치된 OLEDB Provider"
    
    On Error Resume Next
    Dim adoCat As Object
    Set adoCat = CreateObject("ADOX.Catalog")
    
    Dim rsProv As Object
    Set cn = CreateObject("ADODB.Connection")
    Set rsProv = cn.OpenSchema(11) ' adSchemaProviderTypes
    
    If Err.Number = 0 Then
        Do Until rsProv.EOF
            If InStr(LCase(rsProv.fields("TYPE_NAME").value), "ace") > 0 Or _
               InStr(LCase(rsProv.fields("TYPE_NAME").value), "jet") > 0 Or _
               InStr(LCase(rsProv.fields("TYPE_NAME").value), "access") > 0 Then
                Debug.Print "   - " & rsProv.fields("TYPE_NAME").value
            End If
            rsProv.MoveNext
        Loop
        rsProv.Close
    Else
        ' 레지스트리에서 확인 (대체 방법)
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        
        Dim regPaths As Variant
        regPaths = Array( _
            "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Microsoft.ACE.OLEDB.16.0\", _
            "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0\", _
            "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.15.0\", _
            "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0\", _
            "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.Jet.OLEDB.4.0\" _
        )
        
        Dim regPath As Variant
        For Each regPath In regPaths
            Err.Clear
            Dim regValue As String
            regValue = wsh.RegRead(regPath)
            
            If Err.Number = 0 Then
                Debug.Print "   - " & Replace(regPath, "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\", "") & " (레지스트리 발견)"
            End If
        Next
    End If
    
    On Error GoTo 0
    
    ' 9. 진단 결과 요약
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "■ 진단 결과 요약"
    Debug.Print "=========================================="
    
    If successProvider <> "" Then
        Debug.Print "? 연결 가능! 사용할 Provider: " & successProvider
        Debug.Print ""
        Debug.Print "연결 문자열:"
        Debug.Print "Provider=" & successProvider & ";Data Source=" & filePath & ";"
    Else
        Debug.Print "? 연결 불가!"
        Debug.Print ""
        Debug.Print "권장 조치:"
        Debug.Print "1. Access Database Engine 2016 다운로드:"
        
        #If Win64 Then
            Debug.Print "   https://www.microsoft.com/download/details.aspx?id=54920"
            Debug.Print "   → AccessDatabaseEngine_X64.exe 선택"
        #Else
            Debug.Print "   https://www.microsoft.com/download/details.aspx?id=54920"
            Debug.Print "   → AccessDatabaseEngine.exe 선택"
        #End If
        
        Debug.Print ""
        Debug.Print "2. 설치 시 오류가 나면 명령 프롬프트(관리자)에서:"
        
        #If Win64 Then
            Debug.Print "   AccessDatabaseEngine_X64.exe /passive"
        #Else
            Debug.Print "   AccessDatabaseEngine.exe /passive"
        #End If
        
        Debug.Print ""
        Debug.Print "3. 또는 DAO 사용 (도구 > 참조 > Microsoft DAO 3.6 Object Library)"
    End If
    
    Debug.Print ""
    Debug.Print "=========================================="
    
    ' 정리
    Set fso = Nothing
    Set dbFile = Nothing
    
End Sub



