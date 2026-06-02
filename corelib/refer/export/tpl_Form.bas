Attribute VB_Name = "tpl_Form"
Option Explicit
'=====================================================================
' 함수명 : GetAutoColumnWidths
' 설  명 : 배열 데이터의 글자 수를 기반으로 ListBox의 ColumnWidths 값을 자동 계산
' 반환값 : ColumnWidths 속성에 바로 입력 가능한 문자열 (예: "60pt;48pt;72pt")
'---------------------------------------------------------------------
' [매개변수]
'   arr          : ListBox에 바인딩된 2차원 배열 (필수)
'   sngCharWidth : 글자 1개당 너비 (pt 단위, 기본값 = 6)
'                  - 영문/숫자 위주 → 5~6 권장
'                  - 한글 포함     → 8~10 권장
'   sngPadding   : 컬럼 양쪽 여백 (pt 단위, 기본값 = 6)
'                  - 좌/우 각각 적용되므로 실제 총 여백 = sngPadding × 2
'---------------------------------------------------------------------
' [사용 예시]
'   lst.ColumnWidths = GetAutoColumnWidths(arr)          ' 기본값 사용
'   lst.ColumnWidths = GetAutoColumnWidths(arr, 9)       ' 한글 많을 때
'   lst.ColumnWidths = GetAutoColumnWidths(arr, 9, 10)   ' 한글 + 여백 넉넉하게
'=====================================================================
Function GetAutoColumnWidths(arr As Variant, _
                                     Optional sngCharWidth As Single = 6, _
                                     Optional sngPadding As Single = 6) As String
    Dim i As Long, j As Long
    Dim lngMaxLen As Long
    Dim strWidths As String
    
    ' 컬럼(j) 순서대로 순회
    For j = 0 To UBound(arr, 2)
        lngMaxLen = 0
        
        ' 해당 컬럼에서 가장 긴 글자 수 탐색
        For i = 0 To UBound(arr, 1)
            If Len(arr(i, j)) > lngMaxLen Then
                lngMaxLen = Len(arr(i, j))
            End If
        Next i
        
        ' 컬럼 폭 계산: (최대 글자 수 × 글자 너비) + (좌우 여백 합산)
        strWidths = strWidths & (lngMaxLen * sngCharWidth + sngPadding * 2) & "pt;"
    Next j
    
    ' 마지막 세미콜론(;) 제거 후 반환
    GetAutoColumnWidths = Left(strWidths, Len(strWidths) - 1)
End Function
Sub FormSearchFilter(Optional ByVal cbo As Object)
    
    On Error Resume Next
    If cbo Is Nothing Then Set cbo = ActiveSheet.SearchBox
    On Error GoTo 0
    
    If cbo Is Nothing Then Exit Sub
    
    Dim dv As DBvar
    Dim arr As Variant
    
    With dv
        '1. 데이터베이스 ============
        .Alias = [H3]
        .Table = [I3]
        .QPlus = [K3]
        
        If .Table = "" Then
            MsgBox "데이터 테이블을 입력하세요!", vbCritical
            Exit Sub
        End If
        
        If .Alias = "현재워크북" And InStr(.Table, "DB_") = 0 Then  '엑셀 일반 테이블인 경우
            Dim strField As String
            Dim vntValue As Variant
            
            If .QPlus <> "" Then
                strField = Trim(Mid(.QPlus, 1, InStr(.QPlus, "=") - 1))
                vntValue = Trim(Mid(.QPlus, InStr(.QPlus, "=") + 1))
            End If
            '필드
            .fields = [J3]
            arr = GetUniqueValues(.fields, Range(.Table).ListObject, strField, vntValue)
        Else '엑셀 데이터베이스형 테이블인 경우
            
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
            .cstQuery = [L3]
                    
            If .cstQuery <> "" Then '커스텀 쿼리 적용
                .Query = .cstQuery
            Else '자동 생성 쿼리 적용
                
                .Table = IIf(.Type = "엑셀", "[" & .Table & "$]", .Table)
                .fields = [J3]
                .QPlus = [K3]
                
                If .QPlus = "1" Then
                    MsgBox "이 페이지의 검색 값은 필수 입니다.", vbCritical
                    Exit Sub
                End If
                
                .Query = "SELECT DISTINCT " & .fields & " FROM " & .Table & " " & .QPlus
            End If
    
            .arrQuery = Array(.Query)
            
            arr = SelectQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
            arr = arr(0)
        End If
    End With
    
    Call FillComboBox(cbo, arr, cbo.value)
      
End Sub
Function GetUniqueValues(ByVal strField As String, Optional tbl As ListObject, _
                         Optional strFilterField As String = "", _
                         Optional strFilterValue As Variant = "", _
                         Optional returnAs2D As Boolean = False, _
                         Optional useWildcard As Boolean = False) As Variant
    Dim dict As Object
    Dim arrData As Variant
    Dim arrFilter As Variant
    Dim result As Variant
    Dim i As Long, j As Long
    Dim hasFilter As Boolean
    Dim isMatch As Boolean
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' tbl이 입력되지 않으면 활성시트의 첫번째 테이블 지정
    If tbl Is Nothing Then
        If ActiveSheet.ListObjects.count > 0 Then
            Set tbl = ActiveSheet.ListObjects(1)
        Else
            GetUniqueValues = CVErr(xlErrRef)
            Exit Function
        End If
    End If
    
    ' 대상 필드를 배열로 읽기
    On Error Resume Next
    arrData = tbl.ListColumns(strField).DataBodyRange.value
    If Err.Number <> 0 Then
        GetUniqueValues = CVErr(xlErrRef)
        Exit Function
    End If
    On Error GoTo 0
    
    ' 필터 필드가 지정된 경우 해당 열도 배열로 읽기
    hasFilter = (strFilterField <> "")
    If hasFilter Then
        On Error Resume Next
        arrFilter = tbl.ListColumns(strFilterField).DataBodyRange.value
        If Err.Number <> 0 Then
            GetUniqueValues = CVErr(xlErrRef)
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' 배열 순회하며 조건에 맞는 값만 중복 제거
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        ' 필터 조건 확인
        If hasFilter Then
            ' 와일드카드 사용 여부에 따라 비교
            If useWildcard Then
                On Error Resume Next
                isMatch = (arrFilter(i, 1) Like strFilterValue)
                If Err.Number <> 0 Then isMatch = False
                On Error GoTo 0
            Else
                isMatch = (arrFilter(i, 1) = strFilterValue)
            End If
            
            If isMatch Then
                If Not IsEmpty(arrData(i, 1)) And Not dict.Exists(arrData(i, 1)) Then
                    dict.Add arrData(i, 1), ""
                End If
            End If
        Else
            ' 필터가 없으면 모든 값 추가
            If Not IsEmpty(arrData(i, 1)) And Not dict.Exists(arrData(i, 1)) Then
                dict.Add arrData(i, 1), ""
            End If
        End If
    Next i
    
    ' Dictionary의 키를 배열로 변환
    If dict.count > 0 Then
        If returnAs2D Then
            ' 2차원 배열 (세로 출력용)
            ReDim result(1 To dict.count, 1 To 1)
            j = 1
            Dim key As Variant
            For Each key In dict.Keys
                result(j, 1) = key
                j = j + 1
            Next key
        Else
            ' 1차원 배열 (VBA 내부 처리용)
            ReDim result(1 To dict.count)
            j = 1
            For Each key In dict.Keys
                result(j) = key
                j = j + 1
            Next key
        End If
        GetUniqueValues = result
    Else
        GetUniqueValues = CVErr(xlErrNA)
    End If
End Function
Sub FillComboBox(ByVal cbo As Object, ByVal arrValues As Variant, Optional ByVal filterValue As Variant = "")
    ' 콤보박스에 배열 값을 채우는 함수 (필터 기능 추가)
    ' cbo: 콤보박스 객체 (MSForms.ComboBox 또는 Worksheet의 ComboBox)
    ' arrValues: 1차원 배열 또는 2차원 배열
    ' filterValue: 필터링할 값 (이 값이 포함된 항목만 표시)
    
    Dim i As Long
    Dim j As Long
    Dim itemValue As String
    Dim useFilter As Boolean
    
    ' 필터 사용 여부 확인
    useFilter = (filterValue <> "")
    
    ' 콤보박스 초기화
    cbo.Clear
    
    ' 배열이 비어있는지 확인
    If Not IsArray(arrValues) Then
        Exit Sub
    End If
    
    ' 1차원 배열인 경우
    If GetArrayDimensions(arrValues) = 1 Then
        For i = LBound(arrValues) To UBound(arrValues)
            itemValue = CStr(arrValues(i))
            
            ' 필터 조건 확인
            If Not useFilter Then
                ' 필터가 없으면 모두 추가
                cbo.AddItem itemValue
            ElseIf InStr(1, itemValue, CStr(filterValue), vbTextCompare) > 0 Then
                ' 필터 값이 포함된 경우만 추가 (대소문자 구분 안함)
                cbo.AddItem itemValue
            End If
        Next i
        
    ' 2차원 배열인 경우 (첫 번째 차원만 사용)
    ElseIf GetArrayDimensions(arrValues) = 2 Then
        For i = LBound(arrValues, 2) To UBound(arrValues, 2)
            itemValue = CStr(arrValues(0, i))
            
            ' 필터 조건 확인
            If Not useFilter Then
                ' 필터가 없으면 모두 추가
                cbo.AddItem itemValue
            ElseIf InStr(1, itemValue, CStr(filterValue), vbTextCompare) > 0 Then
                ' 필터 값이 포함된 경우만 추가 (대소문자 구분 안함)
                cbo.AddItem itemValue
            End If
        Next i
    End If
    
End Sub
' 배열의 차원 수를 반환하는 헬퍼 함수
Function GetArrayDimensions(arr As Variant) As Integer
    Dim i As Integer
    Dim temp As Long
    
    On Error Resume Next
    i = 0
    Do
        i = i + 1
        temp = UBound(arr, i)
    Loop Until Err.Number <> 0
    
    GetArrayDimensions = i - 1
End Function
Function GetFirstComboBox() As Object
    Dim obj As OLEObject
    
    For Each obj In ActiveSheet.OLEObjects
        If TypeName(obj.Object) = "ComboBox" Then
            Set GetFirstComboBox = obj.Object
            Exit Function
        End If
    Next obj
End Function
Sub SearchBoxResize()
    '<<검색 콤보박스 초기화>>
    On Error Resume Next
    Dim cbo As ComboBox
    Dim rng As Range

    Set cbo = ActiveSheet.SearchBox
    Set rng = Range("C4:D4") ' 크기 기준이 되는 셀 범위

    If IsObject(cbo) Then cbo.value = "" '검색값 지우기
    ChangFormSize cbo, rng '크기 조절
    On Error GoTo 0
End Sub
Sub ChangFormSize(frm As Object, rng As Range)
    ' 폼 컨트롤을 지정된 셀 범위의 크기와 위치에 맞게 조절
    Dim ws As Worksheet
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim dblLeft As Double
    Dim dblTop As Double
    
    ' 워크시트 설정
    Set ws = rng.Worksheet
    
    ' 셀 범위의 크기와 위치 가져오기 (포인트 단위)
    dblWidth = rng.Width
    dblHeight = rng.Height
    dblLeft = rng.Left
    dblTop = rng.Top
    
    ' 폼 컨트롤의 크기와 위치 설정
    With frm
        .Width = dblWidth
        .Height = dblHeight
        .Left = dblLeft
        .Top = dblTop
    End With
End Sub
Sub ClearAllComboBoxes(Optional ByVal sht As Worksheet)
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    Dim obj As Object
    Dim shp As Shape
    
    ' ActiveX 콤보 박스 초기화
    For Each obj In sht.OLEObjects
        If TypeName(obj.Object) = "ComboBox" Then
            obj.Object.Clear
            obj.Object.value = ""
        End If
    Next obj
    
    ' 폼 컨트롤 콤보 박스 초기화
    For Each shp In sht.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlDropDown Then
                shp.ControlFormat.RemoveAllItems
            End If
        End If
    Next shp
End Sub
Sub AddSelectionItems(ByVal rngList As Range, _
                                            Optional ByVal CellExistingValue As Range)
    
    Dim r As Range
    
    With frm_AddList.lst_List
    For Each r In rngList
        .AddItem r
    Next r
    
    Dim arrList() As String
    Dim i, j As Long
    
    arrList = Split(CellExistingValue, ";")
    
    For i = 0 To UBound(arrList)
        For j = 0 To .ListCount - 1
            If .List(j) = arrList(i) Then
                .Selected(j) = True
            End If
        Next
    Next
    End With
    
End Sub
Function GetUnProtectPassword(Optional strPassword As String) As String
    
    With frm_UnProtect
        .txt_Password.value = ""
        If strPassword <> "" Then .lbl_Prompt = strPassword
        .Tag = ""
        .Show
        GetUnProtectPassword = .Tag   ' 함수명과 일치
    End With
    
    Unload frm_UnProtect
    
End Function


