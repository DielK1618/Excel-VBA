Attribute VB_Name = "tpl_Code"
Option Explicit
Sub BtCreateCodes()
    
    Dim strPrefix, strInField As String
    Dim intLen As Integer
    
    strPrefix = [E3]
    intLen = [F3]
    strInField = [G3]
    
    If intLen = 0 Or strInField = "" Then
        MsgBox "코드 입력 필드와 길이 설정을 확인하세요!", vbCritical
        Exit Sub
    End If
    
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes(Application.Caller)
    
    Dim tbl As ListObject
    Set tbl = Range(shp.TopLeftCell, shp.BottomRightCell).ListObject
    
    Dim r, rng As Range
    Set rng = tbl.ListColumns(strInField).DataBodyRange
    
    Dim strToken, strTable, strField As String
    
    strToken = [B3]
    strTable = [C3]
    strField = [D3]
    
    Dim colCodes As New Collection
    
    If strToken <> "" And strTable <> "" And strField <> "" Then
        Set colCodes = GetExistingCodes(strToken, strTable, strField)
    End If
    
    On Error Resume Next
    For Each r In rng
        If r <> "" Then
            colCodes.Add r, r
        End If
    Next
    On Error GoTo 0
    
    cl.DPUpdate_Off
    cl.Calculate_Off
    cl.sht_UnLock
    Set rng = IIf(rng.Rows.count > 1, rng.SpecialCells(xlCellTypeVisible), rng)
    
    For Each r In rng
        If r = "" Then
            r.value = CreateUniqueID(strPrefix, intLen, colCodes)
        End If
    Next
    
    cl.sht_Lock
    cl.Calculate_On
    cl.DPUpdate_On
    
End Sub
Function GetExistingCodes(ByVal strToken As String, _
                                                   ByVal strTable As String, _
                                                   ByVal strField As String) As Collection
                                                   
    On Error GoTo ErrHandler
    Dim dv As DBvar
    
    With dv
        '1. 데이터베이스 ============
        .Token = strToken
        
        If .Token = "" Then
            MsgBox "데이터베이스 예약어를 입력해 주세요!", vbCritical
            Exit Function
        End If
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Token)
        
        .Type = arrDB(0)
        .File = arrDB(3)
        .db = arrDB(4)
        .Server = arrDB(5)
        .Port = arrDB(6)
        .ID = arrDB(7)
        .PW = arrDB(8)
        
        '2. 쿼리 ============
        .Table = strTable
        .Table = IIf(.Type = "엑셀", "[" & .Table & "$]", .Table)
        .fields = strField
        
        .Query = "SELECT DISTINCT " & .fields & " FROM " & .Table
        .arrQuery = Array(.Query)
        .arrData = SelectQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
        
        If IsArrayEmpty(.arrData) Then Exit Function
        
        Dim colTemp As New Collection
        Dim vnt As Variant
        
        On Error Resume Next
        For Each vnt In .arrData(0)
            colTemp.Add vnt, vnt
        Next
        On Error GoTo 0
        
        Set GetExistingCodes = colTemp
        
    End With

Exit Function
ErrHandler:
    Set GetExistingCodes = Nothing
    Debug.Print "오류가 발생했습니다: " & Err.Description, vbCritical
End Function
Function CreateUniqueID(ByVal strPrefix As String, _
                                                ByVal intLen As Integer, _
                                                ByRef idCollection As Collection) As String '약어+숫자
    
    On Error Resume Next
    Dim r As Range
    Dim strID As String
    Dim i As Integer
    Dim attempts As Long
    
    attempts = 0
    
    ' 중복되지 않는 6자리 고유 아이디 생성
    Do
        strID = strPrefix
        ' 랜덤으로 6자리 숫자 생성
        For i = 1 To intLen
            strID = strID & Int((9 * Rnd) + 0)
        Next i
        
        attempts = attempts + 1
        
        ' 안전장치: 너무 오래 걸리면 멈춤 (거의 안 걸림)
        If attempts > 1000 Then
            CreateUniqueID = ""
            Exit Function
        End If
        
    Loop Until CheckUniqueID(idCollection, strID)
    
    idCollection.Add strID, strID
    CreateUniqueID = strID
    
End Function
Function CheckUniqueID(ByVal idCollection As Collection, ByVal strID As String) As Boolean
    
    CheckUniqueID = True
    
    On Error Resume Next
    idCollection.item (strID)
    If Err.Number = 0 Then CheckUniqueID = False
    On Error GoTo 0
    
End Function
Function GenerateRandomCode(ByVal rngCodes As Range, ByVal intLen As Integer) As String '모든 조합 코드
    Dim lowercaseChars As String
    Dim uppercaseChars As String
    Dim numbers As String
    Dim randomChar As String
    Dim i, j, maxAttempts As Long
    Dim ID As String
    Dim r As Range
    Dim idCollection As New Collection
    
    maxAttempts = 1000000
    
    ' 영어 소문자, 대문자, 숫자를 문자열로 정의합니다.
    lowercaseChars = "abcdefghjkmnpqrstuvwxyz"
    uppercaseChars = "ABCDEFGHJKMNPQRSTUVWXYZ"
    numbers = "23456789"
    
    ' 컬렉션에 기존 아이디 담기
    For Each r In rngCodes
        If r <> "" Then idCollection.Add Trim(r), Trim(r)
    Next r
    
    ' 중복되지 않는 아이디 생성
    Do
        ID = ""
        ' 랜덤한 영어 소문자 생성 (intLen-2)자리 생성
        For i = 1 To intLen - 2
            randomChar = Mid(lowercaseChars, Int((Len(lowercaseChars) * Rnd) + 1), 1)
            ID = ID & randomChar
        Next i
        
        ' 랜덤한 영어 대문자 1자리 생성
        randomChar = Mid(uppercaseChars, Int((Len(uppercaseChars) * Rnd) + 1), 1)
        ID = ID & randomChar
        
        ' 랜덤한 숫자 1자리 생성
        randomChar = Mid(numbers, Int((Len(numbers) * Rnd) + 1), 1)
        ID = ID & randomChar
'             최대 시도 횟수 초과 시 루프 종료
        j = j + 1
        If i > maxAttempts Then
            MsgBox "최대 시도 횟수를 초과하여 중복 아이디를 생성할 수 없습니다.", vbExclamation
            Exit Function
        End If
        
    Loop Until CheckUniqueID(idCollection, ID)
    
    Set idCollection = Nothing
    
    ' 최종 아이디 반환
    GenerateRandomCode = ID
End Function
