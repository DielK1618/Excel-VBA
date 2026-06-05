Attribute VB_Name = "am_Utils"
Option Explicit

' +---------------------------------------------------------+
' |  am_Utils                                               |
' |  역할 : 배열/검사/코드생성/날짜/외부앱/도구/수식 범용 유틸리티 |
' +---------------------------------------------------------+

' ==========================================================
'  배열
' ==========================================================

' 목적   : 중첩 배열을 1차원 배열로 평탄화
' 인수   : arr - 중첩 가능한 배열 또는 단일 값
' 반환   : Variant() - 평탄화된 1차원 배열
Public Function ConvertToArrData(ByVal arr As Variant) As Variant
    Dim arrData()  As Variant
    Dim lngIndex   As Long
    Dim lngTotal   As Long

    lngTotal = prv_CountElements(arr)

    If lngTotal > 0 Then
        ReDim arrData(0 To lngTotal - 1)
        prv_FlattenArray arr, arrData, lngIndex
    End If

    ConvertToArrData = arrData
End Function

Private Sub prv_FlattenArray(ByVal arr As Variant, _
                             ByRef arrData() As Variant, _
                             ByRef lngIndex As Long)
    Dim i        As Long
    Dim lngLower As Long
    Dim lngUpper As Long

    If Not IsArray(arr) Then
        arrData(lngIndex) = arr
        lngIndex = lngIndex + 1
        Exit Sub
    End If

    On Error GoTo NotAnArray
    lngLower = LBound(arr)
    lngUpper = UBound(arr)
    On Error GoTo 0

    For i = lngLower To lngUpper
        If IsArray(arr(i)) Then
            prv_FlattenArray arr(i), arrData, lngIndex
        Else
            arrData(lngIndex) = arr(i)
            lngIndex = lngIndex + 1
        End If
    Next i

    Exit Sub

NotAnArray:
    arrData(lngIndex) = arr
    lngIndex = lngIndex + 1
End Sub

Private Function prv_CountElements(ByVal arr As Variant) As Long
    Dim lngCount As Long
    Dim i        As Long
    Dim lngLower As Long
    Dim lngUpper As Long

    If Not IsArray(arr) Then
        prv_CountElements = 1
        Exit Function
    End If

    On Error GoTo NotAnArray
    lngLower = LBound(arr)
    lngUpper = UBound(arr)
    On Error GoTo 0

    For i = lngLower To lngUpper
        If IsArray(arr(i)) Then
            lngCount = lngCount + prv_CountElements(arr(i))
        Else
            lngCount = lngCount + 1
        End If
    Next i

    prv_CountElements = lngCount
    Exit Function

NotAnArray:
    prv_CountElements = 1
End Function

' ==========================================================
'  검사
' ==========================================================

' 목적   : 배열이 비어있는지 확인
' 인수   : arr - 검사할 배열
' 반환   : Boolean - True(비어있음)
Public Function IsArrayEmpty(ByVal arr As Variant) As Boolean
    On Error GoTo ErrHandler
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    Exit Function
ErrHandler:
    IsArrayEmpty = True
End Function

' 목적   : 개체가 Range 타입인지 확인
' 반환   : Boolean
Public Function IsCells(ByVal obj As Object) As Boolean
    IsCells = TypeOf obj Is Range
End Function

' 목적   : Range가 ListObject(표) 내에 있는지 확인
' 인수   : rng          - 검사할 범위
'          tbl          - (반환) 해당 ListObject
'          celTblHeader - (반환) 해당 열의 헤더 셀
' 반환   : Boolean
Public Function IsTableRange(ByVal rng As Range, _
                             Optional ByRef tbl As ListObject, _
                             Optional ByRef celTblHeader As Range) As Boolean
    On Error GoTo isNotTable

    Set tbl = rng.ListObject

    If Not tbl.HeaderRowRange Is Nothing Then
        If rng.Column >= tbl.Range.Column And _
           rng.Column <= tbl.Range.Column + tbl.Range.Columns.Count - 1 Then
            Set celTblHeader = rng.Parent.Cells(tbl.HeaderRowRange.Row, rng.Column)
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

' 목적   : Range 내 셀이 병합되어 있는지 확인
' 인수   : rng - 검사할 범위
' 반환   : Boolean - True(병합 셀 포함)
' 비고   : 원본 tpl_Check 버그 2종 수정
'          1) Selection Is Nothing -> rng Is Nothing
'          2) IsSelectionMerged = True -> IsRangeMerged = True (오타)
Public Function IsRangeMerged(ByVal rng As Range) As Boolean
    Dim cel As Range

    If rng Is Nothing Then
        IsRangeMerged = False
        Exit Function
    End If

    For Each cel In rng.Cells
        If cel.Address <> cel.MergeArea.Address Then
            IsRangeMerged = True
            Exit Function
        End If
    Next cel

    IsRangeMerged = False
End Function

' 목적   : 파일명 유효성 검사 (Windows 기준, 실패 시 MsgBox로 사유 표시)
' 인수   : strFileName - 검사할 파일명 (확장자 포함 가능)
' 반환   : Boolean - True(유효)
Public Function IsValidFileName(ByVal strFileName As String) As Boolean
    Dim strInvalidChars As String
    Dim strBaseName     As String
    Dim i               As Long

    IsValidFileName = False

    If Trim(strFileName) = "" Then
        MsgBox "파일명을 입력하십시오.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    strFileName = Trim(strFileName)

    strInvalidChars = "\/:*?""<>|"
    For i = 1 To Len(strInvalidChars)
        If InStr(strFileName, Mid(strInvalidChars, i, 1)) > 0 Then
            MsgBox "파일명에 사용할 수 없는 문자가 포함되어 있습니다." & vbNewLine & vbNewLine & _
                   "사용 불가 문자: \ / : * ? "" < > |" & vbNewLine & vbNewLine & _
                   "입력된 파일명: " & strFileName, vbExclamation, am_Core.AM_NAME
            Exit Function
        End If
    Next i

    If InStr(strFileName, ".") > 0 Then
        strBaseName = Left(strFileName, InStrRev(strFileName, ".") - 1)
    Else
        strBaseName = strFileName
    End If

    If Trim(strBaseName) = "" Then
        MsgBox "올바른 파일명을 입력하십시오." & vbNewLine & _
               "확장자만 입력할 수 없습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    Select Case UCase(strBaseName)
        Case "CON", "PRN", "AUX", "NUL", _
             "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            MsgBox "'" & strBaseName & "'은(는) Windows 예약 이름으로 사용할 수 없습니다." & vbNewLine & _
                   "다른 파일명을 입력하십시오.", vbExclamation, am_Core.AM_NAME
            Exit Function
    End Select

    If Len(Replace(strBaseName, ".", "")) = 0 Then
        MsgBox "파일명을 점(.)만으로 구성할 수 없습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    If Len(Replace(strBaseName, " ", "")) = 0 Then
        MsgBox "파일명을 공백만으로 구성할 수 없습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    If Len(strFileName) > 255 Then
        MsgBox "파일명이 너무 깁니다. (최대 255자)" & vbNewLine & _
               "현재 길이: " & Len(strFileName) & "자", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    If Right(strFileName, 1) = " " Then
        MsgBox "파일명은 공백으로 끝날 수 없습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    If Right(strFileName, 1) = "." Then
        MsgBox "파일명은 점(.)으로 끝날 수 없습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    IsValidFileName = True
End Function

' 목적   : 셀의 유효성 검사 타입 문자열 반환
' 인수   : cel - 검사할 셀
' 반환   : String - "없음"/"정수"/"소수"/"목록"/"날짜"/"시간"/"텍스트길이"/"사용자정의"
Public Function GetValidationType(ByVal cel As Range) As String
    On Error Resume Next

    Dim lngType As Long
    lngType = cel.Validation.Type

    Select Case lngType
        Case 0: GetValidationType = "없음"
        Case 1: GetValidationType = "정수"
        Case 2: GetValidationType = "소수"
        Case 3: GetValidationType = "목록"
        Case 4: GetValidationType = "날짜"
        Case 5: GetValidationType = "시간"
        Case 6: GetValidationType = "텍스트길이"
        Case 7: GetValidationType = "사용자정의"
    End Select

    If Err.Number <> 0 Then
        GetValidationType = "없음"
        Err.Clear
    End If

    On Error GoTo 0
End Function

' ==========================================================
'  코드 생성
' ==========================================================

' 목적   : 접두사 + 숫자로 구성된 고유 ID 생성
' 인수   : strPrefix    - ID 접두사
'          intLen       - 숫자 부분 자릿수
'          idCollection - 기존 ID 컬렉션 (중복 방지, ByRef로 새 ID 추가됨)
' 반환   : String - 고유 ID ("" 반환 시 1000회 초과로 생성 실패)
Public Function CreateUniqueID(ByVal strPrefix As String, _
                               ByVal intLen As Integer, _
                               ByRef idCollection As Collection) As String
    On Error Resume Next

    Dim strID  As String
    Dim i      As Integer
    Dim lngTry As Long

    Do
        strID = strPrefix
        For i = 1 To intLen
            strID = strID & Int((9 * Rnd) + 0)
        Next i

        lngTry = lngTry + 1
        If lngTry > 1000 Then
            CreateUniqueID = ""
            Exit Function
        End If
    Loop Until prv_CheckUniqueID(idCollection, strID)

    idCollection.Add strID, strID
    CreateUniqueID = strID
End Function

' 목적   : 랜덤 코드 생성 (소문자 + 대문자 1자 + 숫자 1자, 혼동 문자 제외)
' 인수   : rngCodes - 기존 코드 범위 (중복 방지용)
'          intLen   - 코드 전체 길이 (최소 2 이상)
' 반환   : String - 고유 랜덤 코드
Public Function GenerateRandomCode(ByVal rngCodes As Range, _
                                   ByVal intLen As Integer) As String
    Dim strLower   As String
    Dim strUpper   As String
    Dim strNumbers As String
    Dim strChar    As String
    Dim strID      As String
    Dim i          As Long
    Dim j          As Long
    Dim lngMaxTry  As Long
    Dim r          As Range
    Dim colExist   As New Collection

    lngMaxTry = 1000000
    strLower   = "abcdefghjkmnpqrstuvwxyz"
    strUpper   = "ABCDEFGHJKMNPQRSTUVWXYZ"
    strNumbers = "23456789"

    On Error Resume Next
    For Each r In rngCodes
        If r <> "" Then colExist.Add Trim(r), Trim(r)
    Next r
    On Error GoTo 0

    Do
        strID = ""
        For i = 1 To intLen - 2
            strChar = Mid(strLower, Int((Len(strLower) * Rnd) + 1), 1)
            strID = strID & strChar
        Next i

        strChar = Mid(strUpper, Int((Len(strUpper) * Rnd) + 1), 1)
        strID = strID & strChar

        strChar = Mid(strNumbers, Int((Len(strNumbers) * Rnd) + 1), 1)
        strID = strID & strChar

        j = j + 1
        If j > lngMaxTry Then
            MsgBox "최대 시도 횟수를 초과하여 고유 코드를 생성할 수 없습니다.", _
                   vbExclamation, am_Core.AM_NAME
            Exit Function
        End If
    Loop Until prv_CheckUniqueID(colExist, strID)

    Set colExist = Nothing
    GenerateRandomCode = strID
End Function

Private Function prv_CheckUniqueID(ByVal idCollection As Collection, _
                                   ByVal strID As String) As Boolean
    prv_CheckUniqueID = True
    On Error Resume Next
    idCollection.Item strID
    If Err.Number = 0 Then prv_CheckUniqueID = False
    On Error GoTo 0
End Function

' ==========================================================
'  날짜 / 정규식
' ==========================================================

' 목적   : 날짜/시간 문자열을 Excel 시리얼 값으로 변환
' 인수   : strDate - 날짜 문자열 (예: "2024-01-15")
'          strTime - 시간 문자열 (예: "09:30:00", 기본값 "00:00:00")
' 반환   : Double - Excel 시리얼 날짜
Public Function ConvertToExcelSerialDate(ByVal strDate As String, _
                                         Optional ByVal strTime As String = "00:00:00") As Double
    ConvertToExcelSerialDate = DateValue(strDate) + TimeValue(strTime)
End Function

' 목적   : 정규식 패턴으로 문자열에서 캡처 그룹 값 추출
' 인수   : strValue   - 대상 문자열
'          strPattern - 정규식 패턴 (예: "(\d{4})(\d{2})(\d{2})")
' 반환   : Variant() - SubMatches 기준 추출된 값 배열
Public Function ExtractValues(ByVal strValue As String, _
                              ByVal strPattern As String) As Variant
    Dim objRegEx    As Object
    Dim objMatches  As Object
    Dim arrValues() As Variant
    Dim vntSub      As Variant
    Dim i           As Long
    Dim j           As Long

    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Pattern = strPattern
    objRegEx.Global = False

    Set objMatches = objRegEx.Execute(strValue)

    If objMatches.Count > 0 Then
        For i = 0 To objMatches.Count - 1
            For Each vntSub In objMatches(i).SubMatches
                ReDim Preserve arrValues(j)
                arrValues(j) = vntSub
                j = j + 1
            Next vntSub
        Next i
    End If

    ExtractValues = arrValues

    Set objMatches = Nothing
    Set objRegEx = Nothing
End Function

' ==========================================================
'  외부 앱
' ==========================================================

' 목적   : 주소를 Google Maps에서 열기
' 인수   : strAddress - 검색할 주소
Public Sub OpenAddressInGoogleMaps(ByVal strAddress As String)
    Dim strUrl As String
    strAddress = Replace(strAddress, " ", "+")
    strUrl = "https://www.google.com/maps/search/?api=1&query=" & strAddress
    ThisWorkbook.FollowHyperlink strUrl
End Sub

' 목적   : Shell로 동영상 파일의 재생 길이 반환
' 인수   : strPath - 동영상 파일 전체 경로 (mp4/avi/mov/wmv)
' 반환   : String - 재생 길이 문자열 (예: "0:05:23"), 실패 시 ""
Public Function GetVideoLength(ByVal strPath As String) As String
    Dim objShell  As Object
    Dim objFolder As Object
    Dim objFile   As Object

    If Dir(strPath) = "" Then
        MsgBox "파일을 찾을 수 없습니다: " & strPath, vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    Select Case LCase(Right(strPath, 3))
        Case "mp4", "avi", "mov", "wmv"
        Case Else
            MsgBox "지원하지 않는 파일 형식입니다.", vbExclamation, am_Core.AM_NAME
            Exit Function
    End Select

    Set objShell  = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(Left(strPath, InStrRev(strPath, "\")))
    Set objFile   = objFolder.ParseName(Mid(strPath, InStrRev(strPath, "\") + 1))

    GetVideoLength = objFolder.GetDetailsOf(objFile, 27)

    Set objFile   = Nothing
    Set objFolder = Nothing
    Set objShell  = Nothing
End Function

' ==========================================================
'  도구
' ==========================================================

' 목적   : 현재 선택된 개체의 타입을 MsgBox로 표시
Public Sub CheckSelectionType()
    If Selection.Type = xlRange Then
        MsgBox "선택된 개체: 셀(Range)", vbInformation, am_Core.AM_NAME
    ElseIf TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "Shape" Then
        MsgBox "선택된 개체: 도형(Shape)", vbInformation, am_Core.AM_NAME
    Else
        MsgBox "선택된 개체: " & TypeName(Selection), vbInformation, am_Core.AM_NAME
    End If
End Sub

' 목적   : 밀리초 단위 대기 (DoEvents 포함)
' 인수   : lngMs - 대기 시간 (밀리초)
Public Sub WaitMs(ByVal lngMs As Long)
    If lngMs <= 0 Then Exit Sub
    Dim sngStart As Single
    sngStart = Timer
    Do While Timer - sngStart < (lngMs / 1000#)
        DoEvents
    Loop
End Sub

' ==========================================================
'  수식 / 유효성
' ==========================================================

' 목적   : 수식 문자열을 평가하여 Boolean 반환
' 인수   : strFormula - 평가할 수식 문자열 (= 없어도 자동 추가)
' 반환   : Boolean - True(수식 결과가 참)
Public Function EvaluateFormula(ByVal strFormula As String) As Boolean
    On Error GoTo ErrHandler

    Dim vntResult As Variant

    If Left(strFormula, 1) <> "=" Then strFormula = "=" & strFormula

    vntResult = Application.Evaluate(strFormula)

    If IsError(vntResult) Then
        EvaluateFormula = False
    ElseIf VarType(vntResult) = vbBoolean Then
        EvaluateFormula = vntResult
    ElseIf IsNumeric(vntResult) Then
        EvaluateFormula = (vntResult <> 0)
    Else
        EvaluateFormula = False
    End If

    Exit Function

ErrHandler:
    EvaluateFormula = False
End Function

' 목적   : 사용자정의(Type=7) 유효성 수식이 참인 경우에만 셀에 값 설정
' 인수   : rng      - 대상 범위
'          vntValue - 설정할 값
'          blnMsg   - 유효성 실패 시 메시지 표시 여부
' 비고   : 유효성 타입이 7(사용자정의)이 아닌 셀은 유효성 없이 직접 설정
Public Sub SetIfValTrue(ByVal rng As Range, _
                        ByVal vntValue As Variant, _
                        Optional ByVal blnMsg As Boolean = False)
    On Error GoTo ErrHandler

    Dim cel        As Range
    Dim strFormula As String

    For Each cel In rng
        If cel.Validation.Type <> 7 Then GoTo ErrHandler
        strFormula = CStr(cel.Validation.Formula1)
        If EvaluateFormula(strFormula) Then
            cel = vntValue
        ElseIf blnMsg Then
            MsgBox "유효성 검사 실패 — 값을 설정할 수 없습니다.", vbCritical, am_Core.AM_NAME
        End If
    Next cel

    Exit Sub

ErrHandler:
    rng = vntValue
End Sub
