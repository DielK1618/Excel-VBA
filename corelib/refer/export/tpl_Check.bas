Attribute VB_Name = "tpl_Check"
Option Explicit
Function AccessTableExists(strFilePath As String, strTableName As String) As Boolean

    Dim cat As Object
    Dim tbl As Object
    
    On Error GoTo ErrorHandler
    
    AccessTableExists = False
    
    ' ADOX Catalog 생성
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & ";"
    
    ' 테이블 검색
    For Each tbl In cat.Tables
        If tbl.Name = strTableName And tbl.Type = "TABLE" Then
            AccessTableExists = True
            Exit For
        End If
    Next tbl
    
CleanUp:
    Set tbl = Nothing
    Set cat = Nothing
    Exit Function
    
ErrorHandler:
    AccessTableExists = False
    Resume CleanUp
End Function
Function IsTableRange(rng As Range, _
                     Optional ByRef tbl As ListObject, _
                     Optional ByRef celTblHeader As Range) As Boolean
    
    On Error GoTo isNotTable
    Set tbl = rng.ListObject
    
    ' 헤더 행이 있고, 해당 열이 테이블 범위 내에 있는지 확인
    If Not tbl.HeaderRowRange Is Nothing Then
        If rng.Column >= tbl.Range.Column And _
           rng.Column <= tbl.Range.Column + tbl.Range.Columns.count - 1 Then
            Set celTblHeader = Cells(tbl.HeaderRowRange.Row, rng.Column)
        Else
            Set celTblHeader = Nothing
        End If
    Else
        Set celTblHeader = Nothing
    End If
    
    IsTableRange = True
    Exit Function
    
isNotTable:
    Set tbl = Nothing
    Set celTblHeader = Nothing
    
    IsTableRange = False
    
End Function
Function GetValidationType(ByVal cell As Range) As String '데이터 유효성 검사 타입
    On Error Resume Next ' 오류 방지 (유효성 검사가 없는 경우)
    
    Dim validationType As Long
    validationType = cell.Validation.Type
    
    Select Case validationType
        Case 0
            GetValidationType = "없음"
        Case 1
            GetValidationType = "정수"
        Case 2
            GetValidationType = "소수"
        Case 3
            GetValidationType = "목록"
        Case 4
            GetValidationType = "날짜"
        Case 5
            GetValidationType = "시간"
        Case 6
            GetValidationType = "텍스트길이"
        Case 7
            GetValidationType = "사용자지정"
    End Select
    
    If Err.Number <> 0 Then
        GetValidationType = "없음"
        Err.Clear
    End If
    On Error GoTo 0
End Function
Function IsArrayEmpty(arr As Variant) As Boolean
    On Error GoTo ErrHandler
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    Exit Function
ErrHandler:
    IsArrayEmpty = True   ' 오류 발생 시 배열이 비어있다고 처리
End Function
Function IsCells(obj As Object) As Boolean
    IsCells = TypeOf obj Is Range
End Function
Function IsValidFileName(ByVal strFileName As String) As Boolean
    ' ----------------------------------------------------
    ' 파일명 유효성 검사 함수
    ' 입력: 확장자를 포함한 전체 파일명
    ' 반환: True(유효) / False(유효하지 않음)
    ' ----------------------------------------------------
    Dim strInvalidChars As String
    Dim strBaseName As String
    Dim i As Long
    
    ' 초기값
    IsValidFileName = False
    
    ' 1. 빈 문자열 체크
    If Trim(strFileName) = "" Then
        MsgBox "파일명이 비어있습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    strFileName = Trim(strFileName)
    
    ' 2. Windows에서 파일명에 사용할 수 없는 문자 체크
    strInvalidChars = "\/:*?""<>|"
    For i = 1 To Len(strInvalidChars)
        If InStr(strFileName, Mid(strInvalidChars, i, 1)) > 0 Then
            MsgBox "파일명에 사용할 수 없는 문자가 포함되어 있습니다." & vbNewLine & vbNewLine & _
                   "사용 불가 문자: \ / : * ? "" < > |" & vbNewLine & vbNewLine & _
                   "입력한 파일명: " & strFileName, vbExclamation, "유효성 검사 실패"
            Exit Function
        End If
    Next i
    
    ' 3. 확장자 분리 (확장자 제외한 파일명)
    If InStr(strFileName, ".") > 0 Then
        strBaseName = Left(strFileName, InStrRev(strFileName, ".") - 1)
    Else
        strBaseName = strFileName
    End If
    
    ' 4. 확장자 제외한 파일명이 비어있는지 체크 (예: ".accdb"만 입력한 경우)
    If Trim(strBaseName) = "" Then
        MsgBox "올바른 파일명을 입력해주세요." & vbNewLine & _
               "확장자만 입력할 수 없습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 5. Windows 예약 장치명 체크
    Select Case UCase(strBaseName)
        Case "CON", "PRN", "AUX", "NUL", _
             "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            MsgBox "'" & strBaseName & "'은(는) Windows 예약어로 사용할 수 없습니다." & vbNewLine & _
                   "다른 파일명을 입력해주세요.", vbExclamation, "유효성 검사 실패"
            Exit Function
    End Select
    
    ' 6. 파일명이 점으로만 구성되어 있는지 체크 (예: "..." 또는 "..")
    If Len(Replace(strBaseName, ".", "")) = 0 Then
        MsgBox "파일명은 점(.)으로만 구성될 수 없습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 7. 파일명이 공백으로만 구성되어 있는지 체크
    If Len(Replace(strBaseName, " ", "")) = 0 Then
        MsgBox "파일명은 공백으로만 구성될 수 없습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 8. 파일명 길이 체크 (Windows 최대 255자)
    If Len(strFileName) > 255 Then
        MsgBox "파일명이 너무 깁니다. (최대 255자)" & vbNewLine & _
               "현재 길이: " & Len(strFileName) & "자", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 9. 파일명이 공백으로 끝나는지 체크 (Windows에서 문제 발생 가능)
    If Right(strFileName, 1) = " " Then
        MsgBox "파일명은 공백으로 끝날 수 없습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 10. 파일명이 점으로 끝나는지 체크 (확장자가 없는 경우)
    If Right(strFileName, 1) = "." Then
        MsgBox "파일명은 점(.)으로 끝날 수 없습니다.", vbExclamation, "유효성 검사 실패"
        Exit Function
    End If
    
    ' 모든 검사 통과
    IsValidFileName = True
    
End Function
Function IsRangeMerged(ByVal rng As Range) As Boolean
    
    Dim cell As Range
    Dim FirstCell As Range
    
    ' 선택된 범위가 없을 경우 False 반환
    If Selection Is Nothing Then
        IsRangeMerged = False
        Exit Function
    End If
       
    ' 선택된 범위의 각 셀을 순회하며 검사
    For Each cell In rng.Cells
        
        ' 1. 현재 셀의 MergeArea 속성을 가져옵니다.
        ' 2. 현재 셀의 주소와 MergeArea의 주소를 비교합니다.
        '    만약 MergeArea의 주소가 현재 셀의 주소와 다르면 (즉, Cell.Address <> Cell.MergeArea.Address)
        '    이는 현재 셀이 다른 셀들과 병합된 상태라는 의미입니다.
        If cell.Address <> cell.mergeArea.Address Then
            ' 병합된 셀을 발견했으므로 True를 반환하고 함수 종료
            IsSelectionMerged = True
            Exit Function
        End If
        
    Next cell
    
    ' 모든 셀을 검사했지만 병합된 셀이 없으면 False 반환
    IsRangeMerged = False
    
End Function
